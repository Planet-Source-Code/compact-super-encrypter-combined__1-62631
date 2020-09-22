Attribute VB_Name = "Module1"
Function encrypt(strtext As String, strpass As String) As String
Dim crypt As String, buffer As String
Dim i As Integer, ii As Integer, asc1 As Integer, asc2 As Integer, com As Integer

encrypt = ""
For i = 1 To Len(strtext)
ii = ii + 1
asc1 = asc(Mid(strtext, i))
asc1 = Len(CStr(asc1)) + asc1
asc2 = asc(Mid(strpass, ii))
asc2 = Len(CStr(asc2)) + asc2
com = asc1 + asc2
com = Len(CStr(com)) + com + (i - ii)
encrypt = encrypt & com & Chr(1)
If ii >= Len(strpass) Then ii = 0
DoEvents
Next
encrypt = Left$(encrypt, Len(encrypt) - 1)

For i = 1 To Len(encrypt)
If Mid(encrypt, i, 1) <> Chr(1) Then
buffer = buffer & Chr(Mid(encrypt, i, 1) + 147)
Else
buffer = buffer & Chr(1)
End If
DoEvents
Next
encrypt = buffer
DoEvents
End Function

Function decrypt(strtext As String, strpass As String) As String
Dim char() As String, crypt As String, char2() As String
Dim i As Integer, ii As Integer, asc2 As Integer, i1 As Integer

char() = Split(strtext, Chr(1))

For i = 0 To UBound(char)
For ii = 1 To Len(char(i))
crypt = crypt & asc(Mid(char(i), ii)) - 147
Next ii
char(i) = crypt
crypt = ""
Next i
ii = 0
For i = 0 To UBound(char)
i1 = i1 + 1
ii = ii + 1
asc2 = asc(Mid(strpass, ii))
asc2 = Len(CStr(asc2)) + asc2
If Len(CStr(Val(char(i)) - Len(char(i)) - (i1 - ii))) = Len(char(i)) Then
char(i) = char(i) - Len(char(i)) - (i1 - ii)
Else
char(i) = Val(char(i)) - Len(char(i)) + 1 - (i1 - ii)
End If
char(i) = Val(char(i)) - asc2
If Len(CStr(Val(char(i)) - Len(char(i)))) = Len(char(i)) Then
char(i) = Val(char(i)) - Len(char(i))
Else
char(i) = Val(char(i)) - Len(char(i)) + 1
End If
char(i) = Chr(Val(char(i)))
If ii >= Len(strpass) Then ii = 0
If i1 >= Len(strtext) Then i1 = 0
Next
For i = 0 To UBound(char)
crypt = crypt & char(i)
Next
decrypt = crypt
End Function
