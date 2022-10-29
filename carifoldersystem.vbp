Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Function WindowsDir() As String

Dim X As Long
Dim strPath As String
strPath = Space$(1024)
X = GetWindowsDirectory(strPath, Len(strPath))
strPath = Left$(strPath, X)
If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
WindowsDir = strPath
End Function
