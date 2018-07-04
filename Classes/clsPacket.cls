VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------
'-                           -
'-      clsPacket.cls        -
'-                           -
'-  VectorBot Packet Class   -
'-                           -
'-----------------------------

Private sIn As String, sOut As String, sLen As Long

Public Function getPacket()
  getPacket = sIn
End Function

Public Function getOut()
  getOut = sOut
End Function

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

Public Sub sendBNLSPacket(ByVal ID As Byte)
  Dim pLen As String * 2
  
  CopyMemory ByVal pLen, (Len(sOut) + 3), 2
  
  If (frmMain.sckBNLS.State = sckConnected) Then
    frmMain.sckBNLS.SendData pLen & Chr$(ID) & sOut
  End If
End Sub

Public Sub sendPacket(ByVal index As Integer, ByVal ID As Byte)
  Dim pLen As String * 2
  
  CopyMemory ByVal pLen, (Len(sOut) + 4), 2
  
  If frmMain.sckClanMembers(index).State = sckConnected Then
    frmMain.sckClanMembers(index).SendData Chr$(&HFF) & Chr$(ID) & pLen & sOut
  End If
  sOut = vbNullString
End Sub

Public Sub sendChiefPacket(ByVal ID As Byte)
  Dim pLen As String * 2
  
  CopyMemory ByVal pLen, (Len(sOut) + 4), 2
  
  If frmMain.sckChieftain.State = sckConnected Then
    frmMain.sckChieftain.SendData Chr$(&HFF) & Chr$(ID) & pLen & sOut
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
