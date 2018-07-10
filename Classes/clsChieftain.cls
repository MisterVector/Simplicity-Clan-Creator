VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChieftainData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private colFriendsList As New Collection
Private colInitiateList As New Collection
Private chiefQueue As New Collection
Private isReplacingFriendsBool As Boolean
Private colReplaceFriendsList As New Collection

Private Sub Class_Initialize()
    initIdx = 1
End Sub

Public Sub clearFriends()
    For i = 1 To colFriendsList.count
        colFriendsList.Remove 1
    Next i
End Sub

Public Sub addFriend(ByVal chiefFriend As String)
    colFriendsList.Add chiefFriend
End Sub

Public Function getFriend(ByVal idx As Integer)
    getFriend = colFriendsList.Item(idx)
End Function

Public Function getFriendsCount() As Integer
    getFriendsCount = colFriendsList.count
End Function

Public Function getFriendsWithoutInitiatesCount()
    getFriendsWithoutInitiatesCount = UBound(getFriendsExcludeInitiates()) + 1
End Function

Public Function getFriendsExcludeInitiates() As String()
    Dim friends As String

    For Each chiefFriend In colFriendsList
        Dim found As Boolean
    
        For i = 0 To 8
            If LCase(chiefFriend) = LCase(bot(i).username) Then
                found = True
                Exit For
            End If
        Next i
    
        If Not found Then
            If friends <> vbNullString Then
                friends = friends & "|"
            End If
    
            friends = friends & chiefFriend
        End If
    
        found = False
    Next
  
    getFriendsExcludeInitiates = Split(friends, "|")
End Function

Public Function isFriend(ByVal user As String)
    For Each frnd In colFriendsList
        If LCase(user) = LCase(frnd) Then
            isFriend = True
            Exit Function
        End If
    Next
End Function

Public Sub removeFriend(ByVal user As String)
    For i = 1 To colFriendsList.count
        If LCase(colFriendsList.Item(i)) = LCase(user) Then
            colFriendsList.Remove i
            Exit For
        End If
    Next i
End Sub

Public Function getInitiate(ByVal idx As Integer)
    getInitiate = colInitiateList.Item(idx)
End Function

Public Function getInitiateCount()
    getInitiateCount = colInitiateList.count
End Function

Public Function hasInitiate(ByVal init As String) As Boolean
    For Each initiate In colInitiateList
        If LCase(initiate) = LCase(init) Then
            hasInitiate = True
            Exit Function
        End If
    Next
End Function

Public Sub setIsReplacingFriends(ByVal bl As Boolean)
    isReplacingFriendsBool = bl
End Sub

Public Function isReplacingFriends() As Boolean
    isReplacingFriends = isReplacingFriendsBool
End Function

Public Sub addReplaceFriend(ByVal frnd As String)
    colReplaceFriendsList.Add frnd
End Sub

Public Function popReplaceFriend() As String
    popReplaceFriend = colReplaceFriendsList.Item(1)
    colReplaceFriendsList.Remove (1)
End Function

Public Function replaceFriendCount() As Integer
    replaceFriendCount = colReplaceFriendsList.count
End Function

Public Sub addQueue(ByVal str As String)
    chiefQueue.Add str
End Sub

Public Function popQueue() As String
    popQueue = chiefQueue.Item(1)
    chiefQueue.Remove (1)
End Function

Public Function getQueueByArray() As String()
    Dim queueArray() As String
    ReDim queueArray(chiefQueue.count - 1)
  
    For i = 1 To chiefQueue.count
        queueArray(i - 1) = chiefQueue.Item(i)
    Next i
  
    getQueueByArray = queueArray
End Function

Public Sub removeQueueByName(ByVal name As String)
    For i = 1 To chiefQueue.count
        If LCase(chiefQueue.Item(i)) = LCase(name) Then
            chiefQueue.Remove i
            Exit For
        End If
    Next i
End Sub

Public Function isQueueEmpty()
    isQueueEmpty = (chiefQueue.count = 0)
End Function

Public Function isQueueEnabled()
    isQueueEnabled = frmMain.tmrQueue.Enabled
End Function

Public Function getQueueSize() As Integer
    getQueueSize = chiefQueue.count
End Function
