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