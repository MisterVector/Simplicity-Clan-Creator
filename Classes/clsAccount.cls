VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private username As String, password As String

Public Sub initialize()
  
End Sub

Public Sub setUsername(user As String)
  username = user
End Sub

Public Function getUsername() As String
  getUsername = username
End Function

Public Sub setPassword(pass As String)
  password = pass
End Sub

Public Function getPassword() As String
  getPassword = password
End Function


