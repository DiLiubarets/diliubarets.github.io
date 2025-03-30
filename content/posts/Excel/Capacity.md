```vb
Sub CountEmployeesByProject()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pRange As Range
    Dim rCell As Range
    Dim projectCounts As Object
    Dim employeeName As String
    Dim projectName As String
    Dim storyPoints As Double
    
    ' Set the worksheet and pivot table
    Set ws = ActiveSheet ' Change if needed
    Set pt = ws.PivotTables(1) ' Adjust if multiple pivot tables exist
    
    ' Set the data range of the pivot table
    Set pRange = pt.TableRange1
    
    ' Create a dictionary to store project counts
    Set projectCounts = CreateObject("Scripting.Dictionary")
    
    ' Loop through each row in the pivot table
    For Each rCell In pRange.Rows
        ' Check if the row is a subtotal row (usually formatted differently)
        If rCell.Cells(1, 1).Font.Bold Then ' Assuming subtotal rows are bold
            employeeName = rCell.Cells(1, 1).Value
            storyPoints = rCell.Cells(1, 3).Value ' Assuming story points are in column 3
            
            ' If subtotal is greater than 20, count the employee for each project
            If storyPoints > 20 Then
                projectName = rCell.Cells(1, 2).Value ' Assuming project name is in column 2
                
                ' Update project count
                If projectCounts.exists(projectName) Then
                    projectCounts(projectName) = projectCounts(projectName) + 1
                Else
                    projectCounts.Add projectName, 1
                End If
            End If
        End If
    Next rCell
    
    ' Output results
    Dim outputRow As Integer
    outputRow = 2 ' Start output in row 2 of a new sheet
    
    ' Create a new sheet for results
    Dim resultSheet As Worksheet
    On Error Resume Next
    Set resultSheet = ThisWorkbook.Sheets("Results")
    If resultSheet Is Nothing Then
        Set resultSheet = ThisWorkbook.Sheets.Add
        resultSheet.Name = "Results"
    End If
    On Error GoTo 0
    
    ' Clear previous results
    resultSheet.Cells.Clear
    
    ' Write headers
    resultSheet.Cells(1, 1).Value = "Project Name"
    resultSheet.Cells(1, 2).Value = "Employees Over 20 Story Points"
    
    ' Write data
    Dim key As Variant
    For Each key In projectCounts.keys
        resultSheet.Cells(outputRow, 1).Value = key
        resultSheet.Cells(outputRow, 2).Value = projectCounts(key)
        outputRow = outputRow + 1
    Next key
    
    MsgBox "Analysis Complete! Check the 'Results' sheet.", vbInformation
End Sub
```

same sheet 
```vb
Sub CountEmployeesByProject()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pRange As Range
    Dim rCell As Range
    Dim projectCounts As Object
    Dim employeeName As String
    Dim projectName As String
    Dim storyPoints As Variant ' Use Variant to avoid type mismatch errors
    Dim outputRow As Integer
    
    ' Set the worksheet and pivot table
    Set ws = ActiveSheet ' Use the active sheet
    Set pt = ws.PivotTables(1) ' Assuming the first Pivot Table is the target
    
    ' Set the data range of the pivot table
    Set pRange = pt.TableRange1
    
    ' Create a dictionary to store project counts
    Set projectCounts = CreateObject("Scripting.Dictionary")
    
    ' Loop through each row in the pivot table
    For Each rCell In pRange.Rows
        ' Check if the row is a subtotal row (usually formatted differently)
        If rCell.Cells(1, 1).Font.Bold Then ' Assuming subtotal rows are bold
            employeeName = rCell.Cells(1, 1).Value
            storyPoints = rCell.Cells(1, 3).Value ' Assuming story points are in column 3
            
            ' Ensure storyPoints is numeric before comparison
            If IsNumeric(storyPoints) Then
                If storyPoints > 20 Then
                    projectName = rCell.Cells(1, 2).Value ' Assuming project name is in column 2
                    
                    ' Update project count
                    If projectCounts.exists(projectName) Then
                        projectCounts(projectName) = projectCounts(projectName) + 1
                    Else
                        projectCounts.Add projectName, 1
                    End If
                End If
            End If
        End If
    Next rCell
    
    ' Find the row below the Pivot Table to place results
    outputRow = pRange.Rows.Count + pRange.Row + 2 ' Two rows below the Pivot Table
    
    ' Clear previous results (if any)
    ws.Range(ws.Cells(outputRow, 1), ws.Cells(outputRow + 50, 2)).ClearContents
    
    ' Write headers
    ws.Cells(outputRow, 1).Value = "Project Name"
    ws.Cells(outputRow, 2).Value = "Employees Over 20 Story Points"
    
    ' Format headers as bold
    ws.Cells(outputRow, 1).Font.Bold = True
    ws.Cells(outputRow, 2).Font.Bold = True
    
    ' Write data
    Dim key As Variant
    outputRow = outputRow + 1 ' Move to first data row
    For Each key In projectCounts.keys
        ws.Cells(outputRow, 1).Value = key
        ws.Cells(outputRow, 2).Value = projectCounts(key)
        outputRow = outputRow + 1
    Next key
    
    MsgBox "Analysis Complete! Results are placed below the Pivot Table.", vbInformation
End Sub
```

consider project names

```vb
Sub CountEmployeesByProject()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pRange As Range
    Dim rCell As Range
    Dim projectCounts As Object
    Dim employeeProjects As Object
    Dim employeeName As String
    Dim projectName As String
    Dim storyPoints As Variant ' Use Variant to avoid type mismatch errors
    Dim outputRow As Integer
    
    ' Set the worksheet and pivot table
    Set ws = ActiveSheet ' Use the active sheet
    Set pt = ws.PivotTables(1) ' Assuming the first Pivot Table is the target
    
    ' Set the data range of the pivot table
    Set pRange = pt.TableRange1
    
    ' Create dictionaries
    Set projectCounts = CreateObject("Scripting.Dictionary") ' Stores project counts
    Set employeeProjects = CreateObject("Scripting.Dictionary") ' Tracks projects per employee
    
    ' Loop through each row in the pivot table
    For Each rCell In pRange.Rows
        ' Read the first column (Employee Name)
        employeeName = rCell.Cells(1, 1).Value
        projectName = rCell.Cells(1, 2).Value ' Project Name (can be empty for subtotal rows)
        storyPoints = rCell.Cells(1, 3).Value ' Assuming story points are in column 3
        
        ' Check if the row is a subtotal row (bold text)
        If rCell.Cells(1, 1).Font.Bold Then
            ' This is a subtotal row, check if the employee's total is > 20
            If IsNumeric(storyPoints) And storyPoints > 20 Then
                ' If the employee worked on multiple projects, increment count for each
                If employeeProjects.exists(employeeName) Then
                    Dim proj As Variant
                    For Each proj In employeeProjects(employeeName)
                        If projectCounts.exists(proj) Then
                            projectCounts(proj) = projectCounts(proj) + 1
                        Else
                            projectCounts.Add proj, 1
                        End If
                    Next proj
                End If
            End If
            ' Clear employee's project tracking after subtotal
            employeeProjects.Remove employeeName
        ElseIf projectName <> "" Then
            ' This is a normal data row (not a subtotal), track the project for the employee
            If employeeProjects.exists(employeeName) Then
                If Not IsInArray(projectName, employeeProjects(employeeName)) Then
                    employeeProjects(employeeName) = employeeProjects(employeeName) & "," & projectName
                End If
            Else
                employeeProjects.Add employeeName, projectName
            End If
        End If
    Next rCell
    
    ' Find the row below the Pivot Table to place results
    outputRow = pRange.Rows.Count + pRange.Row + 2 ' Two rows below the Pivot Table
    
    ' Clear previous results (if any)
    ws.Range(ws.Cells(outputRow, 1), ws.Cells(outputRow + 50, 2)).ClearContents
    
    ' Write headers
    ws.Cells(outputRow, 1).Value = "Project Name"
    ws.Cells(outputRow, 2).Value = "Employees Over 20 Story Points"
    
    ' Format headers as bold
    ws.Cells(outputRow, 1).Font.Bold = True
    ws.Cells(outputRow, 2).Font.Bold = True
    
    ' Write data
    Dim key As Variant
    outputRow = outputRow + 1 ' Move to first data row
    For Each key In projectCounts.keys
        ws.Cells(outputRow, 1).Value = key
        ws.Cells(outputRow, 2).Value = projectCounts(key)
        outputRow = outputRow + 1
    Next key
    
    MsgBox "Analysis Complete! Results are placed below the Pivot Table.", vbInformation
End Sub

' Helper function to check if a value exists in a comma-separated string
Function IsInArray(value As String, list As String) As Boolean
    Dim items As Variant
    items = Split(list, ",")
    Dim i As Integer
    For i = LBound(items) To UBound(items)
        If Trim(items(i)) = value Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function
```

collection
```vb
Sub CountEmployeesByProject()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pRange As Range
    Dim rCell As Range
    Dim projectCounts As Object
    Dim employeeProjects As Object
    Dim employeeName As String
    Dim projectName As String
    Dim storyPoints As Variant ' Use Variant to avoid type mismatch errors
    Dim outputRow As Integer
    
    ' Set the worksheet and pivot table
    Set ws = ActiveSheet ' Use the active sheet
    Set pt = ws.PivotTables(1) ' Assuming the first Pivot Table is the target
    
    ' Set the data range of the pivot table
    Set pRange = pt.TableRange1
    
    ' Create dictionaries
    Set projectCounts = CreateObject("Scripting.Dictionary") ' Stores project counts
    Set employeeProjects = CreateObject("Scripting.Dictionary") ' Tracks projects per employee
    
    ' Loop through each row in the pivot table
    For Each rCell In pRange.Rows
        ' Read the first column (Employee Name)
        employeeName = rCell.Cells(1, 1).Value
        projectName = rCell.Cells(1, 2).Value ' Project Name (can be empty for subtotal rows)
        storyPoints = rCell.Cells(1, 3).Value ' Assuming story points are in column 3
        
        ' Check if the row is a subtotal row (bold text)
        If rCell.Cells(1, 1).Font.Bold Then
            ' This is a subtotal row, check if the employee's total is > 20
            If IsNumeric(storyPoints) And storyPoints > 20 Then
                ' If the employee worked on multiple projects, increment count for each
                If employeeProjects.exists(employeeName) Then
                    Dim proj As Variant
                    For Each proj In employeeProjects(employeeName)
                        If projectCounts.exists(proj) Then
                            projectCounts(proj) = projectCounts(proj) + 1
                        Else
                            projectCounts.Add proj, 1
                        End If
                    Next proj
                End If
            End If
            ' Remove employee from tracking after subtotal
            If employeeProjects.exists(employeeName) Then
                employeeProjects.Remove employeeName
            End If
        ElseIf projectName <> "" Then
            ' This is a normal data row (not a subtotal), track the project for the employee
            If Not employeeProjects.exists(employeeName) Then
                Dim projCollection As Collection
                Set projCollection = New Collection
                employeeProjects.Add employeeName, projCollection
            End If
            
            ' Add project to employee's collection if not already there
            If Not IsInCollection(employeeProjects(employeeName), projectName) Then
                employeeProjects(employeeName).Add projectName
            End If
        End If
    Next rCell
    
    ' Find the row below the Pivot Table to place results
    outputRow = pRange.Rows.Count + pRange.Row + 2 ' Two rows below the Pivot Table
    
    ' Clear previous results (if any)
    ws.Range(ws.Cells(outputRow, 1), ws.Cells(outputRow + 50, 2)).ClearContents
    
    ' Write headers
    ws.Cells(outputRow, 1).Value = "Project Name"
    ws.Cells(outputRow, 2).Value = "Employees Over 20 Story Points"
    
    ' Format headers as bold
    ws.Cells(outputRow, 1).Font.Bold = True
    ws.Cells(outputRow, 2).Font.Bold = True
    
    ' Write data
    Dim key As Variant
    outputRow = outputRow + 1 ' Move to first data row
    For Each key In projectCounts.keys
        ws.Cells(outputRow, 1).Value = key
        ws.Cells(outputRow, 2).Value = projectCounts(key)
        outputRow = outputRow + 1
    Next key
    
    MsgBox "Analysis Complete! Results are placed below the Pivot Table.", vbInformation
End Sub

' Helper function to check if a value exists in a Collection
Function IsInCollection(col As Collection, value As String) As Boolean
    Dim item As Variant
    On Error Resume Next
    For Each item In col
        If item = value Then
            IsInCollection = True
            Exit Function
        End If
    Next item
    On Error GoTo 0
    IsInCollection = False
End Function
```

hardcode output cell 

```vb
Sub CountEmployeesByProject()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pRange As Range
    Dim rCell As Range
    Dim projectCounts As Object
    Dim employeeProjects As Object
    Dim employeeName As String
    Dim projectName As String
    Dim storyPoints As Variant ' Use Variant to avoid type mismatch errors
    Dim outputRow As Integer
    
    ' Set the worksheet and pivot table
    Set ws = ActiveSheet ' Use the active sheet
    Set pt = ws.PivotTables(1) ' Assuming the first Pivot Table is the target
    
    ' Set the data range of the pivot table
    Set pRange = pt.TableRange1
    
    ' Create dictionaries
    Set projectCounts = CreateObject("Scripting.Dictionary") ' Stores project counts
    Set employeeProjects = CreateObject("Scripting.Dictionary") ' Tracks projects per employee
    
    ' Loop through each row in the pivot table
    For Each rCell In pRange.Rows
        ' Read the first column (Employee Name)
        employeeName = rCell.Cells(1, 1).Value
        projectName = rCell.Cells(1, 2).Value ' Project Name (can be empty for subtotal rows)
        storyPoints = rCell.Cells(1, 3).Value ' Assuming story points are in column 3
        
        ' Check if the row is a subtotal row (bold text)
        If rCell.Cells(1, 1).Font.Bold Then
            ' This is a subtotal row, check if the employee's total is > 20
            If IsNumeric(storyPoints) And storyPoints > 20 Then
                ' If the employee worked on multiple projects, increment count for each
                If employeeProjects.exists(employeeName) Then
                    Dim proj As Variant
                    For Each proj In employeeProjects(employeeName)
                        If projectCounts.exists(proj) Then
                            projectCounts(proj) = projectCounts(proj) + 1
                        Else
                            projectCounts.Add proj, 1
                        End If
                    Next proj
                End If
            End If
            ' Remove employee from tracking after subtotal
            If employeeProjects.exists(employeeName) Then
                employeeProjects.Remove employeeName
            End If
        ElseIf projectName <> "" Then
            ' This is a normal data row (not a subtotal), track the project for the employee
            If Not employeeProjects.exists(employeeName) Then
                Dim projCollection As Collection
                Set projCollection = New Collection
                employeeProjects.Add employeeName, projCollection
            End If
            
            ' Add project to employee's collection if not already there
            If Not IsInCollection(employeeProjects(employeeName), projectName) Then
                employeeProjects(employeeName).Add projectName
            End If
        End If
    Next rCell
    
    ' Set output location to J1
    Dim outputCol As String
    outputCol = "J"
    outputRow = 1  ' Start from row 1
    
    ' Clear previous results (if any)
    ws.Range(ws.Cells(outputRow, 10), ws.Cells(outputRow + 50, 11)).ClearContents ' J and K columns
    
    ' Write headers
    ws.Cells(outputRow, 10).Value = "Project Name" ' Column J
    ws.Cells(outputRow, 11).Value = "Employees Over 20 Story Points" ' Column K
    
    ' Format headers as bold
    ws.Cells(outputRow, 10).Font.Bold = True
    ws.Cells(outputRow, 11).Font.Bold = True
    
    ' Write data
    Dim key As Variant
    outputRow = outputRow + 1 ' Move to first data row
    For Each key In projectCounts.keys
        ws.Cells(outputRow, 10).Value = key ' Column J
        ws.Cells(outputRow, 11).Value = projectCounts(key) ' Column K
        outputRow = outputRow + 1
    Next key
    
    MsgBox "Analysis Complete! Results are placed in columns J and K.", vbInformation
End Sub

' Helper function to check if a value exists in a Collection
Function IsInCollection(col As Collection, value As String) As Boolean
    Dim item As Variant
    On Error Resume Next
    For Each item In col
        If item = value Then
            IsInCollection = True
            Exit Function
        End If
    Next item
    On Error GoTo 0
    IsInCollection = False
End Function
```

```vb
Sub CountEmployeesByProject()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pRange As Range
    Dim rCell As Range
    Dim projectCounts As Object
    Dim employeeProjects As Object
    Dim employeeName As String
    Dim projectName As String
    Dim storyPoints As Variant ' Use Variant to avoid type mismatch errors
    Dim outputRow As Integer
    Dim dataRange As Range
    
    ' Set the worksheet and pivot table
    Set ws = ActiveSheet ' Use the active sheet
    Set pt = ws.PivotTables(1) ' Assuming the first Pivot Table is the target
    
    ' Set the data range of the pivot table
    Set pRange = pt.TableRange1
    
    ' Create dictionaries
    Set projectCounts = CreateObject("Scripting.Dictionary") ' Stores project counts
    Set employeeProjects = CreateObject("Scripting.Dictionary") ' Tracks projects per employee
    
    ' Loop through each row in the pivot table
    For Each rCell In pRange.Rows
        ' Read the first column (Employee Name)
        employeeName = rCell.Cells(1, 1).Value
        projectName = rCell.Cells(1, 2).Value ' Project Name (can be empty for subtotal rows)
        storyPoints = rCell.Cells(1, 3).Value ' Assuming story points are in column 3
        
        ' Check if the row is a subtotal row (bold text)
        If rCell.Cells(1, 1).Font.Bold Then
            ' Remove " Total" from the employee name
            employeeName = Replace(employeeName, " Total", "")
            
            ' This is a subtotal row, check if the employee's total is > 20
            If IsNumeric(storyPoints) And storyPoints > 20 Then
                ' If the employee worked on multiple projects, increment count for each
                If employeeProjects.exists(employeeName) Then
                    Dim proj As Variant
                    For Each proj In employeeProjects(employeeName)
                        If projectCounts.exists(proj) Then
                            projectCounts(proj) = projectCounts(proj) + 1
                        Else
                            projectCounts.Add proj, 1
                        End If
                    Next proj
                End If
            End If
            ' Remove employee from tracking after subtotal
            If employeeProjects.exists(employeeName) Then
                employeeProjects.Remove employeeName
            End If
        ElseIf projectName <> "" Then
            ' This is a normal data row (not a subtotal), track the project for the employee
            If Not employeeProjects.exists(employeeName) Then
                Dim projCollection As Collection
                Set projCollection = New Collection
                employeeProjects.Add employeeName, projCollection
            End If
            
            ' Add project to employee's collection if not already there
            If Not IsInCollection(employeeProjects(employeeName), projectName) Then
                employeeProjects(employeeName).Add projectName
            End If
        End If
    Next rCell
    
    ' Set output location to J1
    Dim outputCol As String
    outputCol = "J"
    outputRow = 1  ' Start from row 1
    
    ' Clear previous results (if any)
    ws.Range(ws.Cells(outputRow, 10), ws.Cells(outputRow + 50, 11)).ClearContents ' J and K columns
    
    ' Write headers
    ws.Cells(outputRow, 10).Value = "Project Name" ' Column J
    ws.Cells(outputRow, 11).Value = "Employees Over 20 Story Points" ' Column K
    
    ' Format headers as bold
    ws.Cells(outputRow, 10).Font.Bold = True
    ws.Cells(outputRow, 11).Font.Bold = True
    
    ' Write data
    Dim key As Variant
    outputRow = outputRow + 1 ' Move to first data row
    For Each key In projectCounts.keys
        ws.Cells(outputRow, 10).Value = key ' Column J
        ws.Cells(outputRow, 11).Value = projectCounts(key) ' Column K
        outputRow = outputRow + 1
    Next key
    
    ' Apply conditional formatting to column K
    Set dataRange = ws.Range(ws.Cells(2, 11), ws.Cells(outputRow - 1, 11)) ' K2 to last row
    
    ' Clear previous conditional formatting
    dataRange.FormatConditions.Delete
    
    ' Apply red format (4 or more employees)
    With dataRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="4")
        .Interior.Color = RGB(255, 0, 0) ' Red
        .Font.Color = RGB(255, 255, 255) ' White text
    End With
    
    ' Apply yellow format (1 or 2 employees)
    With dataRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLessEqual, Formula1:="2")
        .Interior.Color = RGB(255, 255, 0) ' Yellow
        .Font.Color = RGB(0, 0, 0) ' Black text
    End With
    
    ' Apply green format (exactly 3 employees)
    With dataRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="3")
        .Interior.Color = RGB(0, 255, 0) ' Green
        .Font.Color = RGB(0, 0, 0) ' Black text
    End With
    
    MsgBox "Analysis Complete! Results are placed in columns J and K with conditional formatting.", vbInformation
End Sub

' Helper function to check if a value exists in a Collection
Function IsInCollection(col As Collection, value As String) As Boolean
    Dim item As Variant
    On Error Resume Next
    For Each item In col
        If item = value Then
            IsInCollection = True
            Exit Function
        End If
    Next item
    On Error GoTo 0
    IsInCollection = False
End Function
```

```vb
Sub CopyPivotTableData()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim cell As Range
    Dim copyRange As Range
    Dim destCell As Range
    Dim lastCol As Integer
    Dim headerRows As Integer
    Dim rowRange As Range
    Dim pasteRange As Range
    
    ' Set the worksheet and pivot table
    Set ws = ThisWorkbook.Sheets("Summary")
    Set pt = ws.PivotTables("GereralPivotTable")
    
    ' Set the destination cell (starting point for pasting)
    Set destCell = ws.Range("S4")
    
    ' Define the number of header rows
    headerRows = 2 ' Since your Pivot Table has two header rows
    
    ' Identify the last column in the pivot table (Grand Total column is usually the last one)
    lastCol = pt.DataBodyRange.Columns.Count
    
    ' Copy headers (first two rows) from the Pivot Table and paste into the destination
    pt.TableRange1.Rows(1).Resize(headerRows).Copy
    destCell.PasteSpecial Paste:=xlPasteValues
    destCell.PasteSpecial Paste:=xlPasteFormats
    
    ' Loop through the Grand Total column to find rows where the value is >21
    For Each cell In pt.DataBodyRange.Columns(lastCol).Cells
        ' Ensure the row is NOT a "Grand Total" row
        If IsNumeric(cell.Value) And cell.Value > 21 Then
            ' Check if the row contains "Grand Total" in the first column (adjust if needed)
            If LCase(cell.EntireRow.Cells(1, 1).Value) <> "grand total" Then
                ' Get the actual row range within the entire Pivot Table (including row fields)
                Set rowRange = Intersect(cell.EntireRow, pt.TableRange1)
                
                ' If first row to copy, set copyRange
                If copyRange Is Nothing Then
                    Set copyRange = rowRange
                Else
                    ' Extend the range to include this row
                    Set copyRange = Union(copyRange, rowRange)
                End If
            End If
        End If
    Next cell
    
    ' Copy and paste the filtered rows
    If Not copyRange Is Nothing Then
        copyRange.Copy
        Set pasteRange = destCell.Offset(headerRows, 0)
        pasteRange.PasteSpecial Paste:=xlPasteValues
        pasteRange.PasteSpecial Paste:=xlPasteFormats
        
        ' AutoFit the pasted table
        ws.Range(pasteRange, pasteRange.End(xlToRight)).Columns.AutoFit
    End If
    
    ' Clean up
    Application.CutCopyMode = False
    MsgBox "Data copied successfully and columns adjusted!", vbInformation

End Sub
```

```vb
Function CountSelectedItems(slicerName As String) As Integer
    Dim sl As Slicer
    Dim si As SlicerItem
    Dim count As Integer
    
    Set sl = ThisWorkbook.SlicerCaches(slicerName).Slicers(1)
    count = 0
    
    For Each si In sl.SlicerCache.SlicerItems
        If si.Selected Then
            count = count + 1
        End If
    Next si
    
    CountSelectedItems = count
End Function
=CountSelectedItems("Slicer_Region")
```