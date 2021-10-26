Attribute VB_Name = "Protect_WS"
' UnProtect All Worksheets
Sub UnProtectAllSheets()
Attribute UnProtectAllSheets.VB_ProcData.VB_Invoke_Func = "P\n14"
    Dim ws As Worksheet
    Dim pw As String
    pw = "?"


    For Each ws In Worksheets
        ws.Unprotect 'pw    ' uncomment to apply password
    Next ws

End Sub


' Protect All Worksheets
Sub ProtectAllSheets()
Attribute ProtectAllSheets.VB_ProcData.VB_Invoke_Func = "p\n14"
    Dim ws As Worksheet
    Dim pw As String
    pw = "?"

    For Each ws In Worksheets
        ws.Protect 'pw  ' uncomment to apply password
        
    Next ws

End Sub


Private Sub Crack_Password()
''' Crack password on the sheet


'''''''''' Variable Declarations
Dim try_pw As String

Dim i As Integer, j As Integer, k As Integer
Dim l As Integer, m As Integer, n As Integer
Dim i1 As Integer, i2 As Integer, i3 As Integer
Dim i4 As Integer, i5 As Integer, i6 As Integer


'''''''''' For Loops
On Error Resume Next
For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126

'''''''''' Try password
try_pw = Chr(i) & Chr(j) & Chr(k) & _
    Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
    Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)

ActiveSheet.Unprotect try_pw

'''''''''' State
If ActiveSheet.ProtectContents = False Then
    MsgBox try_pw

ActiveWorkbook.Sheets(1).Select
Range("a1").FormulaR1C1 = try_pw
    Exit Sub


End If
Next: Next: Next: Next: Next: Next
Next: Next: Next: Next: Next: Next
End Sub

