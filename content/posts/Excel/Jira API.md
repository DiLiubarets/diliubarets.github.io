
[Releases · VBA-tools/VBA-JSON](https://github.com/VBA-tools/VBA-JSON/releases)

```vb
Sub GetAPIData()
    Dim http As Object
    Dim JSON As Object
    Dim url As String
    Dim username As String
    Dim password As String
    Dim ws As Worksheet
    Dim rowNum As Integer, colNum As Integer
    Dim firstRecord As Object
    Dim key As Variant
    Dim headers As Object
    
    ' API Credentials (Modify as needed)
    username = "your-email@example.com"
    password = "your-api-token"

    ' API URL (Modify for your needs)
    url = "https://your-jira-instance.atlassian.net/rest/api/3/search?jql=project=MYPROJECT"

    ' Create HTTP Request Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False, username, password
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Basic " & Base64Encode(username & ":" & password)
    http.Send
    
    ' Parse JSON Response
    Set JSON = JsonConverter.ParseJson(http.responseText)
    
    ' Get first record to determine headers dynamically
    If JSON("issues").Count = 0 Then
        MsgBox "No data found!", vbExclamation
        Exit Sub
    End If
    
    ' Select worksheet and clear old data
    Set ws = ThisWorkbook.Sheets("API_Data")
    ws.Cells.ClearContents
    
    ' Extract headers from first record
    Set firstRecord = JSON("issues")(1)("fields") ' Adjust based on API structure
    colNum = 1
    Set headers = CreateObject("Scripting.Dictionary")

    For Each key In firstRecord.keys
        headers.Add key, colNum
        ws.Cells(1, colNum).Value = key  ' Write headers
        colNum = colNum + 1
    Next key

    ' Loop through records and write data
    rowNum = 2
    For Each issue In JSON("issues")
        colNum = 1
        For Each key In firstRecord.keys
            If issue("fields").Exists(key) Then
                ws.Cells(rowNum, headers(key)).Value = issue("fields")(key)
            Else
                ws.Cells(rowNum, headers(key)).Value = "N/A"
            End If
        Next key
        rowNum = rowNum + 1
    Next issue

    MsgBox "Data Imported Successfully!", vbInformation
End Sub

' Function to encode Base64 for authentication
Function Base64Encode(text As String) As String
    Dim arr() As Byte
    Dim objXML As Object
    Dim objNode As Object
    
    arr = StrConv(text, vbFromUnicode)
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arr
    Base64Encode = objNode.Text
End Function

```

safe loop
```vb
For Each issue In JSON("issues")
    For Each key In firstRecord.keys
        If issue("fields").Exists(key) Then
            On Error Resume Next
            If IsObject(issue("fields")(key)) Then
                ws.Cells(rowNum, headers(key)).Value = "[Object]"
            Else
                ws.Cells(rowNum, headers(key)).Value = issue("fields")(key)
            End If
            On Error GoTo 0
        Else
            ws.Cells(rowNum, headers(key)).Value = "N/A"
        End If
    Next key
    rowNum = rowNum + 1
Next issue
```