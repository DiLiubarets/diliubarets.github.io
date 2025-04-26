```vb
Private Sub Workbook_SlicerChange(ByVal Slicer As Slicer)
    Dim slicer1 As SlicerCache
    Dim slicer2 As SlicerCache
    Dim item As SlicerItem
    Dim selectedItems As Collection
    Dim i As Long
    Dim sourceSlicerName As String
    Dim targetSlicerName As String
    
    ' Define your slicer names
    Dim slicerName1 As String
    Dim slicerName2 As String
    
    slicerName1 = "Slicer_SlicerName1" ' First slicer name
    slicerName2 = "Slicer_SlicerName2" ' Second slicer name
    
    ' Determine which slicer was changed
    If Slicer.SlicerCache.Name = slicerName1 Then
        sourceSlicerName = slicerName1
        targetSlicerName = slicerName2
    ElseIf Slicer.SlicerCache.Name = slicerName2 Then
        sourceSlicerName = slicerName2
        targetSlicerName = slicerName1
    Else
        Exit Sub ' Not one of the slicers we want to sync
    End If
    
    ' Set slicer caches
    Set slicer1 = ThisWorkbook.SlicerCaches(sourceSlicerName)
    Set slicer2 = ThisWorkbook.SlicerCaches(targetSlicerName)
    
    ' Create a collection to hold selected items
    Set selectedItems = New Collection
    
    ' Loop through source slicer items and store selected items
    For Each item In slicer1.SlicerItems
        If item.Selected Then
            selectedItems.Add item.Name
        End If
    Next item
    
    ' Prevent error if slicer is cleared
    On Error Resume Next
    slicer2.ClearManualFilter
    On Error GoTo 0
    
    ' Deselect all items first
    For Each item In slicer2.SlicerItems
        item.Selected = False
    Next item
    
    ' Select items in target slicer that match source slicer
    For i = 1 To selectedItems.Count
        On Error Resume Next
        slicer2.SlicerItems(selectedItems(i)).Selected = True
        On Error GoTo 0
    Next i
End Sub

```