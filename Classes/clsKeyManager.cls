VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private dicKeys As Dictionary
Private keyIdx As Long

Private Sub Class_Initialize()
  keyIdx = 0
  Set dicKeys = New Dictionary
End Sub

Public Sub addKey(ByVal key As String)
  dicKeys.Add dicKeys.count, key
End Sub

Public Sub removeKey(ByVal idx As Long)
  dicKeys.Item(idx) = vbNullString
End Sub

Public Function getKeyByIndex(ByVal idx As Long) As String
  getKeyByIndex = dicKeys.Item(idx)
End Function

Public Function getKey(ByRef keyIndex As Long) As String
  If (keyIdx <= dicKeys.count) Then
    getKey = dicKeys.Item(keyIdx)
    keyIndex = keyIdx
    keyIdx = keyIdx + 1
  End If
End Function

Public Sub resetKeys()
  keyIdx = 0
  dicKeys.RemoveAll
End Sub

Public Function getCount() As Long
  getCount = dicKeys.count
End Function
