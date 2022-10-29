Private Function Encrypt(ByVal astring)

Dim x, i, tmp
For i = 1 To Len(astring)
    x = Mid(astring, i, 1)
tmp = tmp & Chr(Asc(x) + 1)
Next

tmp = StrReverse(tmp)
Encrypt = tmp
End Function

Private Function Decrypt(ByVal encryptedstring)

Dim x, i, tmp
    encryptedstring = StrReverse(encryptedstring)
For i = 1 To Len(encryptedstring)
    x = Mid(encryptedstring, i, 1)
    tmp = tmp & Chr(Asc(x) - 1)
Next
    Decrypt = tmp
End Function

Private Sub Command1_Click()
    Text2.Text = Encrypt(Text1.Text)
End Sub

Private Sub Command2_Click()
    Text3.Text = Decrypt(Text2.Text)
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
  Text1.Text = ""
  Text2.Text = ""
  Text3.Text = ""
End Sub

Private Sub Form_Load()
  Text1.Text = ""
End Sub
