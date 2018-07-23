Attribute VB_Name = "modFileHandling"
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Sub WriteINI(ByVal cSection As String, ByVal cKey As String, ByVal cValue As String, ByVal cPath As String)
    cPath = App.path & "\" & cPath
  
    WritePrivateProfileString cSection, cKey, cValue, cPath
End Sub

Public Function ReadINI(ByVal cSection As String, ByVal cKey As String, ByVal cPath As String) As String
    Dim cBuff As String, cLen As Long
  
    cPath = App.path & "\" & cPath
    cBuff = String$(255, vbNull)
    cLen = GetPrivateProfileString(cSection, cKey, Chr$(0), cBuff, 255, cPath)
  
    If (cLen > 0) Then
        ReadINI = Left$(cBuff, cLen) 'Split(cBuff, Chr$(0))(0)
    Else
        ReadINI = vbNullString
    End If
End Function
