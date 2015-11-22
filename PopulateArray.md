##Populate an array
Populate a one-dimensional array from a range in the workbook.
```vba
Sub PopulateArray(vArray As Variant, rSource As Range)
    Dim j As Integer
        For j = 1 To Selection.Rows.Count
            vArray(j) = Selection.Cells(j).Value
        Next j
End Sub
```
