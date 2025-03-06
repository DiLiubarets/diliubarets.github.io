```vb
Sub FilterPivotTable()
    Dim pt As PivotTable
    Set pt = ActiveSheet.PivotTables(1) ' Adjust if needed

    With pt.PivotFields("Sum of Story Points")
        .ClearAllFilters
        .PivotFilters.Add2 Type:=xlValueIsGreaterThan, Value1:=20
    End With
End Sub
```
```vb
Sub AddSlicerToPivotTable()

    Dim ws As Worksheet
    Dim pt As Excel.PivotTable
    Dim slicerCache As Excel.SlicerCache
    Dim slicer As Excel.Slicer

    ' Set worksheet reference
    Set ws = ThisWorkbook.Sheets("Executive_Report")

    ' Set PivotTable reference
    Set pt = ws.PivotTables("Executive_Pivot")

    ' Create the slicer cache
    Set slicerCache = ThisWorkbook.SlicerCaches.Add(pt, "Program Name")

    ' Add the slicer to the worksheet
    Set slicer = slicerCache.Slicers.Add(ws, , "Slicer_ProgramName", "Program Name", 0, 100, 0, 200)

    ' Style the slicer
    With slicer
        .Style = "SlicerStyleOther1" ' Change to your preferred style
        .Height = 57
        .Width = 296
        .Top = 27
        .Left = 50
        .NumberOfColumns = 4 ' Set the number of columns to 4
        .DisplayHeader = False
    End With

End Sub
```