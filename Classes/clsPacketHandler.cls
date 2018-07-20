VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacketHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sIn As String, sOut As String, sLen As Long
Private sck As Winsock, pt As packetType

Public Function getPacket()
    getPacket = sIn
End Function

Public Function getOut()
    getOut = sOut
End Function

Public Sub setup(socket As Winsock, packetType As packetType)
    Set sck = socket
    pt = packetType
End Sub

Public Sub SetData(ByVal sData As String)
    sIn = sData
    sLen = 1
End Sub

Public Sub clearIn()
    sIn = vbNullString
End Sub

Public Sub clearOut()
    sOut = vbNullString
End Sub

Public Sub Skip(Bytes As Long)
    sLen = sLen + Bytes
End Sub

Public Sub InsertByte(ByVal value As Byte)
    sOut = sOut & Chr$(value)
End Sub

Public Sub InsertWORD(ByVal value As Integer)
    Dim NewValue As String * 2
    CopyMemory ByVal NewValue, value, 2
    sOut = sOut & NewValue
End Sub

Public Sub InsertDWORD(ByVal value As Long)
    Dim NewValue As String * 4
    CopyMemory ByVal NewValue, value, 4
    sOut = sOut & NewValue
End Sub

Public Sub InsertNTString(ByVal value As String)
    sOut = sOut & value & vbNullChar
End Sub

Public Sub InsertNonNTString(ByVal value As String)
    sOut = sOut & value
End Sub

Public Sub sendPacket(ByVal ID As Byte)
    Dim pLen As String * 2, outLength As Long, header As String
    
    outLength = (Len(sOut) + IIf(pt = packetType.BNCS, 4, 3))
    
    CopyMemory ByVal pLen, outLength, 2
    
    header = IIf(pt = packetType.BNCS, Chr$(&HFF) & Chr$(ID) & pLen, pLen & Chr$(ID))
    
    If (sck.State = sckConnected) Then
        sck.SendData header & sOut
    End If
    
    sOut = vbNullString
End Sub

Public Function GetByte() As Byte
    GetByte = Asc(Mid$(sIn, sLen, 1))
    sLen = sLen + 1
End Function

Public Function GetWORD() As Long
    Dim NewValue As Long
    CopyMemory NewValue, ByVal Mid$(sIn, sLen, 2), 2
    GetWORD = NewValue
    sLen = sLen + 2
End Function

Public Function GetDWORD() As Long
    Dim NewValue As Long
    CopyMemory NewValue, ByVal Mid$(sIn, sLen, 4), 4
    GetDWORD = NewValue
    sLen = sLen + 4
End Function

Public Function GetNonNTString(ByVal index As Integer) As String
    GetNonNTString = Mid$(sIn, sLen, index)
    sLen = sLen + Len(GetNonNTString)
End Function

Public Function getNTString() As String
    getNTString = Split(Mid$(sIn, sLen), vbNullChar)(0)
    sLen = sLen + Len(getNTString) + 1
End Function
