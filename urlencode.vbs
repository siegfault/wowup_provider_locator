' https://gist.github.com/greenkey/dbb317fbf229d4f85e0f

' Encode special characters of a string
' this is useful when you want to put a string in the URL
' inspired by http://stackoverflow.com/questions/218181/how-can-i-url-encode-a-string-in-excel-vba
Public Function URLEncode( StringVal )
  Dim i, CharCode, Char, Space
  Dim StringLen

  StringLen = Len(StringVal)
  ReDim result(StringLen)

  Space = "+"
  'Space = "%20"

  For i = 1 To StringLen
    Char = Mid(StringVal, i, 1)
    CharCode = AscW(Char)
    If 97 <= CharCode And CharCode <= 122 _
    Or 64 <= CharCode And CharCode <= 90 _
    Or 48 <= CharCode And CharCode <= 57 _
    Or 45 = CharCode _
    Or 46 = CharCode _
    Or 95 = CharCode _
    Or 126 = CharCode Then
      result(i) = Char
    ElseIf 32 = CharCode Then
      result(i) = Space
    Else
      result(i) = "&#" & CharCode & ";"
    End If
  Next
  URLEncode = Join(result, "")
End Function
