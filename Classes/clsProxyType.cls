VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProxyType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private IP As String
Private Port As Long
Private Version As String

Public Sub setIP(s As String)
  IP = s
End Sub

Public Function getIP() As String
  getIP = IP
End Function

Public Function setPort(p As Long)
  Port = p
End Function

Public Function getPort() As Long
  getPort = Port
End Function

Public Sub setVersion(s As String)
  Version = s
End Sub

Public Function getVersion() As String
  getVersion = Version
End Function
