```vba
Sub toCSV()
    Dim r As Range
    Dim myFile As String
    Dim i, j As Integer
    Dim cellValue As Variant
    
    Set r = Selection
    'MsgBox r.Columns.Count
    'MsgBox r.Rows.Count
    
    MsgBox Application.ActiveWorkbook.Path                      ' Get current file's path
    myFile = Application.ActiveWorkbook.Path & "\output.csv"    ' Create a new output file
    
    Open myFile For Output As #1
    For i = 1 To r.Rows.Count
        For j = 1 To r.Columns.Count
            cellValue = r.Cells(i, j).Value
            If j = r.Columns.Count Then
                Write #1, cellValue
            Else
                Write #1, cellValue,
            End If
        Next j
    Next i
    Close #1
End Sub
```
