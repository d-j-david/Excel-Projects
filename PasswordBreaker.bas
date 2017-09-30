Attribute VB_Name = "PasswordBreaker"
Sub BreakPassword()
' =================================================================================================
' The original author is unfortunately lost to internet antiquity. I simply cleaned up the code,
' and am now storing it for easier access in the future.
'
' NOTE: This approach is only effective against worksheets protected in Excel 2010 or earlier.
'       For Excel 2013 on, protection is still breakable, but requires a more rigorous approach.
' =================================================================================================

    Dim i0 As Integer, i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer, i5 As Integer
    Dim i6 As Integer, i7 As Integer, i8 As Integer, i9 As Integer, iA As Integer, iB As Integer
    Dim i As Long

    MsgBox "This could take a few seconds to a few minutes." & vbCrLf & vbCrLf & _
           "Progress will be shown on the status bar."

    On Error Resume Next
    i = 1
    For i0 = 65 To 66: For i1 = 65 To 66: For i2 = 65 To 66: For i3 = 65 To 66
    For i4 = 65 To 66: For i5 = 65 To 66: For i6 = 65 To 66: For i7 = 65 To 66
    For i8 = 65 To 66: For i9 = 65 To 66: For iA = 65 To 66: For iB = 32 To 126

        ActiveSheet.Unprotect Chr(i0) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & _
                              Chr(i6) & Chr(i7) & Chr(i8) & Chr(i9) & Chr(iA) & Chr(iB)

        ' Give status update
        Application.StatusBar = "Attempt " & i & " of 194560 possible attempts."
        i = i + 1
        DoEvents

        If ActiveSheet.ProtectContents = False Then
            
            MsgBox "One usable password is: " & vbCrLf & vbCrLf & _
                    Chr(i0) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & _
                    Chr(i6) & Chr(i7) & Chr(i8) & Chr(i9) & Chr(iA) & Chr(iB)
            
            Application.StatusBar = False
            Exit Sub
        End If

    Next: Next: Next: Next: Next: Next: Next: Next: Next: Next: Next: Next

End Sub

