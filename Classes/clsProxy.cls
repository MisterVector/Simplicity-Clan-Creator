VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type ProxyData
  IP As String
  Port As Long
  Version As String
End Type

Private arrProxies() As ProxyData
Private pIdx As Long
Private ppIdx As Integer
Private connectionsPerProxy As Integer

Private Sub Class_Initialize()
  ReDim arrProxies(0)
  pIdx = 0
  ppIdx = 0
End Sub

Public Sub setConnectionsPerProxy(perProxy As Integer)
  connectionsPerProxy = perProxy
End Sub

Public Sub addProxy(IP As String, Port As Long, Version As String)
  If Not arrProxies(0).IP = vbNullString Then ReDim Preserve arrProxies(UBound(arrProxies) + 1)
  
  With arrProxies(UBound(arrProxies))
    .IP = IP
    .Port = Port
    .Version = Version
  End With
End Sub

Public Sub clearProxies()
  ReDim arrProxies(0)
End Sub

Public Function getProxy() As clsProxyType
  Set getProxy = New clsProxyType
    
  If pIdx <= UBound(arrProxies) Then
    With getProxy
      Call .setIP(IIf(pIdx <= UBound(arrProxies), arrProxies(pIdx).IP, ""))
      Call .setPort(IIf(pIdx <= UBound(arrProxies), arrProxies(pIdx).Port, ""))
      Call .setVersion(IIf(pIdx <= UBound(arrProxies), arrProxies(pIdx).Version, ""))
    End With
    
    ppIdx = ppIdx + 1
    
    If ppIdx = connectionsPerProxy Then
      ppIdx = 0
      pIdx = pIdx + 1
    End If
  End If
End Function

Public Sub resetProxies()
  pIdx = 0
  ppIdx = 0
End Sub
