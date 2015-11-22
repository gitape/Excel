```vba
Function WeekDay(i As Integer) As String
      Dim s As String
      Select Case i
          Case 1: s = "Monday"
          Case 2: s = "Tuesday"
          Case 3: s = "Wednesday"
          Case 4: s = "Thursday"
          Case 5: s = "Friday"
          Case 6: s = "Saturday"
          Case Else
          s = "Sunday"
      End Select
      WeekDay = s
End Function
```
