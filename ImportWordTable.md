```vba
Sub ImportWordTable()
	Dim wdDoc As Object
	Dim wdFileName As Variant
	Dim tableNo As Integer 'table number in Word
	Dim iRow As Long 'row index in Excel
	Dim iCol As Integer 'column index in Excel
	Dim resultRow As Long
	Dim tableStart As Integer
	Dim tableTot As Integer
	On Error Resume Next

	ActiveSheet.Range("A:AZ").ClearContents
	
	wdFileName = Application.GetOpenFilename("Word files (*.docx),*.docx", , _
	"Browse for file containing table to be imported")
	
	If wdFileName = False Then Exit Sub '(user cancelled import file browser)
	
	Set wdDoc = GetObject(wdFileName) 'open Word file
	
	With wdDoc
	    tableNo = wdDoc.tables.Count
	    tableTot = wdDoc.tables.Count
	    If tableNo = 0 Then
	        MsgBox "This document contains no tables", _
	        vbExclamation, "Import Word Table"
	    ElseIf tableNo > 1 Then
	        tableNo = InputBox("This Word document contains " & tableNo & " tables." & vbCrLf & _
	        "Enter the table to start from", "Import Word Table", "1")
	    End If
	
	    resultRow = 4
	        For tableStart = 1 To tableTot
	        With .tables(tableStart)
	            'copy cell contents from Word table cells to Excel cells
	            For iRow = 1 To .Rows.Count
	                For iCol = 1 To .Columns.Count
	                    Cells(resultRow, iCol) = WorksheetFunction.Clean(.cell(iRow, iCol).Range.Text)
	                Next iCol
	                resultRow = resultRow + 1
	            Next iRow
	        End With
	        resultRow = resultRow + 1
	    Next tableStart
	End With
End Sub
```
