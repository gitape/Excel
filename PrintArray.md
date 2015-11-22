## Print an Array

```vba
Sub PrintArray(vArray As Variant, rDestination As Range)
    Dim i As Integer
    For i = 1 To UBound(vArray)
        rDestination.Offset(i - 1, 0).Value = vArray(i)
    Next
End Sub
```
