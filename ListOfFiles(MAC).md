### List Of files
Written on Mac with Office 2011

```vba
Function ListOfFiles(sFolderPath As String) As Variant
        
    Dim vListOfFiles() As Variant
    Dim strFile As String
    Dim i As Integer
    
    strFile = Dir(sFolderPath)
        
    If Len(strFile) > 0 Then
        i = 1
        ReDim vListOfFiles(i)
        vListOfFiles(i) = strFile
    End If
    
    'Loop through each file in the folder
    Do While Len(strFile) > 0
        strFile = Dir
        If Len(strFile) > 0 Then
            i = i + 1
            ReDim Preserve vListOfFiles(i)
            vListOfFiles(i) = strFile
        End If
    Loop
    
    ListOfFiles = vListOfFiles

End Function
```
