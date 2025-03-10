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