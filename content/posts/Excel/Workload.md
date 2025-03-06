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