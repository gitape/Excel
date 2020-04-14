```vba
Function findDiscRate(r As Range) As Variant
    Dim dateRange As Range
    Dim rateRange As Range
    Dim result As Variant
    Dim lookupValue As Variant
    Dim i, j As Integer
    Dim xrow, xcol As Integer
    
    lookupValue = r.Value
    Set dateRange = Worksheets("Discount Rates").Range("AW2:CE70")
    Set rateRange = Worksheets("Discount Rates").Range("N2:AU70")
    xrow = 0
    xcol = 0
    
    For i = 1 To dateRange.Rows.Count
        For j = 1 To dateRange.Columns.Count
            If dateRange.Cells(i, j).Value = lookupValue Then
                xrow = i
                xcol = j
            End If
        Next j
    Next i
    
    'Debug.Print lookupValue & " x = " & xrow & " y = " & xcol
    
    result = rateRange.Cells(xrow, xcol).Value
    
    findDiscRate = result

End Function
```
