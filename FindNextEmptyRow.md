### Find next empty row
Start with a cell and move down until a non-empty cell is found.

```vba
Function FindNextEmptyRow(r As Range) As Range
    Do While r.Value <> ""
        Set r = r.Offset(1, 0)
    Loop
    Set FindNextEmptyRow = r
End Function
```
