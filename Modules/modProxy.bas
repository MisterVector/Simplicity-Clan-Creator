Attribute VB_Name = "modProxy"
Public Type ProxyData
    IP As String
    Port As Long
    Version As String
End Type

Private arrProxies() As ProxyData
Private pIdx As Long
Private ppIdx As Integer

Public Sub addProxy(IP As String, Port As Long, Version As String)
    If (Not arrProxies(0).IP = vbNullString) Then ReDim Preserve arrProxies(UBound(arrProxies) + 1)
  
    With arrProxies(UBound(arrProxies))
        .IP = IP
        .Port = Port
        .Version = Version
    End With
End Sub

Public Function getProxy() As ProxyData
    Dim proxy As ProxyData
    
    If (pIdx <= UBound(arrProxies)) Then
        proxy = arrProxies(pIdx)
  
        ppIdx = ppIdx + 1
    
        If (ppIdx = config.connectsPerProxy) Then
            ppIdx = 0
            pIdx = pIdx + 1
        End If
    End If
  
    getProxy = proxy
End Function

Public Function countProxies() As Integer
    Dim count As Long

    For i = 0 To UBound(arrProxies)
        If (arrProxies(i).IP <> vbNullString) Then
            count = count + 1
        End If
    Next i
  
    countProxies = count
End Function

Public Sub resetProxies()
    pIdx = 0
    ppIdx = 0
    ReDim arrProxies(0)
End Sub
