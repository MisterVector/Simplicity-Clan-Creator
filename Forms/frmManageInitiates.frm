VERSION 5.00
Begin VB.Form frmManageInitiates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Initiates"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3255
   Icon            =   "frmManageInitiates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnUpdate 
      Caption         =   "Update Password"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   3015
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton btnRemove 
      Caption         =   "Remove"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   3015
   End
   Begin VB.ListBox lstInitiates 
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "Password"
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
      Left            =   1680
      TabIndex        =   11
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Username"
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
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Double-clicking an initiate in the list will let you update the password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "If this list runs out, Simplicity will switch to its ordinary behavior with generating initiate names."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Custom Initiate List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmManageInitiates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clickedUsername As String

Private Sub btnClose_Click()
  Unload Me
End Sub

Private Sub btnAdd_Click()
  Dim str As String
  
  If txtUsername.text = vbNullString Or txtPassword.text = vbNullString Then
    MsgBox "Username and Password must not be blank.", vbOKOnly & vbInformation, PROGRAM_TITLE
    Exit Sub
  End If
  
  If Len(txtUsername.text) < 3 Then
    MsgBox "Username must be 3 characters or more.", vbOKOnly & vbInformation, PROGRAM_TITLE
    Exit Sub
  End If
  
  If Not initiateManager.addCustomInitiate(txtUsername.text, txtPassword.text) Then
    MsgBox "That initiate is already on the initiate list!", vbOKOnly & vbExclamation, PROGRAM_TITLE
    Exit Sub
  End If
  
  lstInitiates.AddItem txtUsername.text
  
  txtUsername.text = vbNullString
  txtPassword.text = vbNullString
End Sub

Private Sub btnRemove_Click()
  Dim username As String, initiateArray() As String

  If lstInitiates.List(lstInitiates.ListIndex) = vbNullString Then Exit Sub
  username = lstInitiates.List(lstInitiates.ListIndex)
  lstInitiates.RemoveItem lstInitiates.ListIndex
  
  initiateManager.removeInitiate username
  
  txtUsername.text = ""
  txtPassword.text = ""
  
  btnUpdate.Enabled = False
  btnAdd.Enabled = True
End Sub

Private Sub btnUpdate_Click()
  initiateManager.updateInitiate txtUsername.text, txtPassword.text
  
  btnUpdate.Enabled = False
  btnAdd.Enabled = True
  
  txtUsername.text = ""
  txtPassword.text = ""
End Sub

Private Sub Form_Load()
  Dim accounts() As New clsAccount
  accounts = initiateManager.getInitiates()
  
  For i = 0 To UBound(accounts)
    If accounts(i).getUsername() <> "" Then
      lstInitiates.AddItem accounts(i).getUsername()
    End If
  Next i
End Sub

Private Sub lstInitiates_Click()
  Dim username As String
  username = lstInitiates.List(lstInitiates.ListIndex)
  
  If username = clickedUsername Then
    Exit Sub
  End If

  btnUpdate.Enabled = False
  btnAdd.Enabled = True
  
  txtUsername.text = ""
  txtPassword.text = ""
End Sub

Private Sub lstInitiates_DblClick()
  Dim username As String
  username = lstInitiates.List(lstInitiates.ListIndex)
  
  If username = "" Then
    Exit Sub
  End If
  
  clickedUsername = username
  btnUpdate.Enabled = True
  btnAdd.Enabled = False
  
  Dim acc As clsAccount
  Set acc = initiateManager.getAccountByName(username)
  
  txtUsername.text = acc.getUsername()
  txtPassword.text = acc.getPassword()
  
  txtPassword.SelStart = 0
  txtPassword.SelLength = Len(txtPassword.text)
  txtPassword.SetFocus
End Sub

Private Sub txtUsername_Click()
  btnUpdate.Enabled = False
  btnAdd.Enabled = True
End Sub

