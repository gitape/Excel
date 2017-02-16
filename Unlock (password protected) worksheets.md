### Unlock password protected files
From: http://jsbi.blogspot.com/2008/09/how-to-easily-unprotectremove-password.html

```vba
Sub UnlockWorkbook()
'
' Breaks worksheet and workbook structure passwords. Jason S
' probably originator of base code algorithm modified for coverage
' of workbook structure / windows passwords and for multiple passwords
' Jason S http://jsbi.blogspot.com
' Reveals hashed passwords NOT original passwords
Const DBLSPACE As String = vbNewLine & vbNewLine
Const AUTHORS As String = DBLSPACE & vbNewLine & "Adapted from Bob McCormick base code by" & "Jason S http://jsbi.blogspot.com"
Const HEADER As String = "AllInternalPasswords User Message"
Const VERSION As String = DBLSPACE & "Version 1.0 8 Sep 2008"
Const REPBACK As String = DBLSPACE & "Please report failure to jasonblr@gmail.com "
Const ALLCLEAR As String = DBLSPACE & "The workbook should be cleared"
Const MSGNOPWORDS1 As String = "There were no passwords on " & AUTHORS & VERSION
Const MSGNOPWORDS2 As String = "There was no protection to " & "workbook structure or windows." & DBLSPACE

Const MSTAKETIME As String = "After pressing OK button this " & _
"will take some time." & DBLSPACE & "Amount of time" & _
"depends on how many different passwords, the"

Const MSGPWORDFOUND1 As String = "You had a Worksheet " & "Structure or Windows Password set." _
& DBLSPACE & "The password found was: " & DBLSPACE & "$$" & DBLSPACE & _
"Note it down for potential future use in other workbooks by " & "the same person who set this password." _
& DBLSPACE & "Now to check and clear other passwords." & AUTHORS & VERSION

Const MSGPWORDFOUND2 As String = "You had a Worksheet " & "password set." & DBLSPACE & "The password found was: " & DBLSPACE & "$$" & DBLSPACE & _
"Note it down for potential " & "future use in other workbooks by same person who " & _
"set this password." & DBLSPACE & "Now to check and clear " & "other passwords." & AUTHORS & VERSION

Const MSGONLYONE As String = "Only structure / windows " & "protected with the password that was just found." & ALLCLEAR & AUTHORS & VERSION & REPBACK
Dim w1 As Worksheet, w2 As Worksheet
Dim i As Integer, j As Integer, k As Integer, l As Integer
Dim m As Integer, n As Integer, i1 As Integer, i2 As Integer
Dim i3 As Integer, i4 As Integer, i5 As Integer, i6 As Integer
Dim PWord1 As String
Dim ShTag As Boolean, WinTag As Boolean
Application.ScreenUpdating = False
With ActiveWorkbook
WinTag = .ProtectStructure Or .ProtectWindows

End With
ShTag = False
For Each w1 In Worksheets
ShTag = ShTag Or w1.ProtectContents
Next w1
If Not ShTag And Not WinTag Then
MsgBox MSGNOPWORDS1, vbInformation, HEADER
Exit Sub
End If
MsgBox MSGTAKETIME, vbInformation, HEADER
If Not WinTag Then
    MsgBox MSGNOPWORDS2, vbInformation, HEADER
Else

On Error Resume Next
Do 'dummy do loop
For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
With ActiveWorkbook
    .Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
    Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    
    If .ProtectStructure = False And .ProtectWindows = False Then
        PWord1 = Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
        Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
        MsgBox Application.Substitute(MSGPWORDFOUND1, "$$", PWord1), vbInformation, _
        HEADER
        Exit Do 'Bypass all for...nexts
    End If
End With
Next: Next: Next: Next: Next: Next
Next: Next: Next: Next: Next: Next
Loop Until True
On Error GoTo 0
End If
If WinTag And Not ShTag Then
    MsgBox MSGONLYONE, vbInformation, HEADER
    Exit Sub
End If

On Error Resume Next
For Each w1 In Worksheets
'Attempt clearance with PWord1
    w1.Unprotect PWord1
Next w1
On Error GoTo 0
ShTag = False

For Each w1 In Worksheets
    'Checks for all clear ShTag triggered to 1 if not.
    ShTag = ShTag Or w1.ProtectContents
Next w1

If ShTag Then
    For Each w1 In Worksheets
        With w1
        If .ProtectContents Then
            On Error Resume Next
            Do 'Dummy do loop
                For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
                For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
                For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
                For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
                .Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
                Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                If Not .ProtectContents Then
                    PWord1 = Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
                    Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                    MsgBox Application.Substitute(MSGPWORDFOUND2, "$$", PWord1), vbInformation, HEADER
                    'leverage finding Pword by trying on other sheets
                    For Each w2 In Worksheets
                        w2.Unprotect PWord1
                    Next w2
                    Exit Do 'Bypass all for...nexts
                End If
                Next: Next: Next: Next: Next: Next
                Next: Next: Next: Next: Next: Next
            Loop Until True
            On Error GoTo 0
        End If
        End With
    Next w1
End If
MsgBox ALLCLEAR & AUTHORS & VERSION & REPBACK, vbInformation, HEADER
End Sub

```
