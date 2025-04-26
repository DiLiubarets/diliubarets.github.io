```vb
Private Sub Workbook_SlicerChange(ByVal Slicer As Slicer)
    Dim sourceSlicerName As String
    Dim targetSlicerName As String
    Dim slicerName1 As String
    Dim slicerName2 As String
    Dim slicer1 As SlicerCache
    Dim slicer2 As SlicerCache
    Dim item As SlicerItem
    Dim selectedItems As Collection
    Dim i As Long
    
    ' Define your slicer names
    slicerName1 = "Slicer_SlicerName1" ' First slicer
    slicerName2 = "Slicer_SlicerName2" ' Second slicer
    
    ' Check if Slicer object is valid
    If Slicer Is Nothing Then Exit Sub
    If Slicer.SlicerCache Is Nothing Then Exit Sub
    
    ' Determine which slicer was changed
    If Slicer.SlicerCache.Name = slicerName1 Then
        sourceSlicerName = slicerName1
        targetSlicerName = slicerName2
    ElseIf Slicer.SlicerCache.Name = slicerName2 Then
        sourceSlicerName = slicerName2
        targetSlicerName = slicerName1
    Else
        Exit Sub ' Not a slicer we want to sync
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
    
    ' Clear filters on target slicer
    On Error Resume Next
    slicer2.ClearManualFilter
    On Error GoTo 0
    
    ' Deselect all first
    For Each item In slicer2.SlicerItems
        item.Selected = False
    Next item
    
    ' Select matching items
    For i = 1 To selectedItems.Count
        On Error Resume Next
        slicer2.SlicerItems(selectedItems(i)).Selected = True
        On Error GoTo 0
    Next i
End Sub

```