VERSION 5.00
Begin VB.Form frmFriendsList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chieftain's Friends List"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3465
   Icon            =   "frmFriendsList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   3465
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrEnableButton 
      Interval        =   1750
      Left            =   7320
      Top             =   2040
   End
   Begin VB.CommandButton btnContinue 
      Caption         =   "Continue"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdRemoveFriend 
      Caption         =   "Remove Friend"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.ListBox lstFriendsList 
      Height          =   2010
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblFriendsToRemove 
      Alignment       =   2  'Center
      Caption         =   "Friends to remove: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   3135
   End
End
Attribute VB_Name = "frmFriendsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public friendsToRemove As Integer
Public cleanClose As Boolean

Public Sub addFriends(ByVal remain As Integer)
    Dim friendsExcludingInitiates() As String
    friendsExcludingInitiates = chiefData.getFriendsExcludeInitiates()
    
    For i = 0 To UBound(friendsExcludingInitiates)
        lstFriendsList.AddItem friendsExcludingInitiates(i)
    Next i
    
    friendsToRemove = remain
    lblFriendsToRemove.Caption = "Friends to remove: " & friendsToRemove
End Sub

Private Sub btnContinue_Click()
    Dim frnd As String
    
    chiefData.setIsReplacingFriends True
    
    AddChat vbYellow, "Simplicity will now add the rest of the initiates."
    frmMain.tmrQueue.Enabled = True
    
    cleanClose = True
    Unload Me
End Sub

Private Sub cmdRemoveFriend_Click()
    Dim frnd As String
    
    frnd = lstFriendsList.List(lstFriendsList.ListIndex)
    If frnd = vbNullString Then Exit Sub
    
    With chiefPacket
        .InsertNTString "/friends remove " & frnd
        lstFriendsList.RemoveItem (lstFriendsList.ListIndex)
        .sendChiefPacket &HE
    End With
    chiefData.removeFriend frnd
  
    cmdRemoveFriend.Enabled = False
    friendsToRemove = friendsToRemove - 1
    lblFriendsToRemove.Caption = "Friends to remove: " & friendsToRemove
  
    If friendsToRemove = 0 Then
        btnContinue.Enabled = True
    Else
        tmrEnableButton.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    cleanClose = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cleanClose Then Exit Sub

    Dim msgBoxResult As Integer
    msgBoxResult = MsgBox("Are you sure you want to do that?", vbYesNo Or vbQuestion, PROGRAM_TITLE)
  
    If msgBoxResult = vbNo Then
        Cancel = 1
    Else
        resetAll
    End If
End Sub

Private Sub tmrEnableButton_Timer()
    cmdRemoveFriend.Enabled = True
    tmrEnableButton.Enabled = False
End Sub
