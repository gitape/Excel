### Linear interpolation function
VERY rough draft, doesn't work in edge cases.

```vba
Function lint(x As Range, y As Range, z As Range) As Double
    Dim i, j, k As Integer
    Dim low_index, high_index As Integer
    Dim rise_over_run As Double
    Dim d As Double
    
    If x.Rows.Count <> y.Rows.Count Then
        MsgBox ("x and y are not compatible")
        lint = 0
    Else
        i = 1
        
        Do While Int(CDbl(x(i + 1).Value)) < Int(CDbl(z.Value))
            i = i + 1
        Loop
       
        low_index = i
        high_index = i + 1
        
        rise_over_run = (y(high_index).Value - y(low_index).Value) / (Int(CDbl(x(high_index).Value)) - Int(CDbl(x(low_index).Value)))
        
        d = y(low_index) + rise_over_run * (Int(CDbl(z.Value)) - Int(CDbl(x(low_index).Value)))
        
        lint = d
    
    End If
End Function
```
