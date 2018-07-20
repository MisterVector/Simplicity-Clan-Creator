VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simplicity Clan Creator v%v"
   ClientHeight    =   7455
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   8145
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckBNLS 
      Left            =   6840
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrChiefReconnect 
      Enabled         =   0   'False
      Left            =   6840
      Top             =   5280
   End
   Begin VB.Timer tmrChiefTimeout 
      Enabled         =   0   'False
      Left            =   5880
      Top             =   5760
   End
   Begin MSWinsockLib.Winsock sckChieftain 
      Left            =   6360
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdResetProxies 
      Caption         =   "Reset Proxies"
      Height          =   480
      Left            =   6720
      TabIndex        =   22
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Timer tmrCheckUpdate 
      Enabled         =   0   'False
      Interval        =   450
      Left            =   5400
      Top             =   5760
   End
   Begin MSWinsockLib.Winsock sckUpdateCheck 
      Left            =   5400
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clan Configuration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   29
      Top             =   1920
      Width           =   7935
      Begin VB.CommandButton btnManageInitiates 
         Caption         =   "Manage Initiates"
         Height          =   375
         Left            =   6240
         TabIndex        =   15
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CheckBox chkCustomInitiates 
         Caption         =   "Use Custom Initiates"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4080
         TabIndex        =   14
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CheckBox chkSaveClanInfo 
         Caption         =   "Save clan configuration to file on exit"
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
         TabIndex        =   13
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox txtInitiatesPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   6000
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtInitiate 
         Height          =   285
         Left            =   6000
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtClanDescription 
         Height          =   315
         Left            =   2160
         MaxLength       =   64
         TabIndex        =   10
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtChief 
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtChiefPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtClanTag 
         Height          =   285
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblInitiateDescription 
         Caption         =   "The initiate username will have the numbers 1-9 appended to the end of each name."
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
         Left            =   4080
         TabIndex        =   36
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label Label8 
         Caption         =   "Initiate Password:"
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
         Left            =   4080
         TabIndex        =   35
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblPrefix 
         Caption         =   "Initiate Username:"
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
         Left            =   4080
         TabIndex        =   34
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Clan Description:"
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
         TabIndex        =   33
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Chieftain's Username:"
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
         TabIndex        =   32
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Chieftain's password:"
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
         TabIndex        =   31
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Clan Tag: "
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
         TabIndex        =   30
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.CommandButton btnRefreshKeys 
      Caption         =   "Reload Keys"
      Height          =   480
      Left            =   5400
      TabIndex        =   21
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Timer tmrReconnect 
      Enabled         =   0   'False
      Index           =   0
      Left            =   6360
      Top             =   5280
   End
   Begin VB.Timer tmrInitiateTimeout 
      Enabled         =   0   'False
      Index           =   0
      Left            =   5880
      Top             =   5280
   End
   Begin MSWinsockLib.Winsock sckClanMembers 
      Index           =   0
      Left            =   5880
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrQueue 
      Enabled         =   0   'False
      Interval        =   1750
      Left            =   5400
      Top             =   5280
   End
   Begin VB.Frame Frame1 
      Caption         =   "Main Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   7935
      Begin VB.CheckBox chkCheckUpdateOnStartup 
         Caption         =   "Check for Update on Startup"
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
         Left            =   4080
         TabIndex        =   39
         Top             =   1515
         Width           =   2895
      End
      Begin VB.CheckBox chkRememberWindowPosition 
         Caption         =   "Remember Window Position"
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
         TabIndex        =   6
         Top             =   1515
         Width           =   3015
      End
      Begin VB.TextBox txtBNLSServer 
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtReconnectTime 
         Height          =   285
         Left            =   6000
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtPerProxy 
         Height          =   285
         Left            =   6000
         TabIndex        =   5
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtTimeOut 
         Height          =   285
         Left            =   6000
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtChannel 
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
      End
      Begin VB.ComboBox cmbServer 
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "BNLS Server"
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
         TabIndex        =   38
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Reconnect Time:"
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
         Left            =   4080
         TabIndex        =   28
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Connects Per Proxy:"
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
         Left            =   4080
         TabIndex        =   27
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Connection Timeout"
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
         Left            =   4080
         TabIndex        =   26
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Channel To Join"
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
         TabIndex        =   25
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Battle.Net Gateway:"
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
         TabIndex        =   24
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton btnRefresh 
      Caption         =   "Disconnect All"
      Enabled         =   0   'False
      Height          =   480
      Left            =   4080
      TabIndex        =   20
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton btnCheckClanned 
      Caption         =   "Check Keys"
      Height          =   480
      Left            =   2640
      TabIndex        =   19
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton btnCheckClanTag 
      Caption         =   "Check Clan Tag"
      Enabled         =   0   'False
      Height          =   480
      Left            =   1320
      TabIndex        =   18
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton btnCreateClan 
      Caption         =   "Create Clan"
      Enabled         =   0   'False
      Height          =   480
      Left            =   120
      TabIndex        =   17
      Top             =   6480
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   2055
      Left            =   120
      TabIndex        =   16
      Top             =   4320
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   3625
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0CCA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblConnected 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Connected:  0"
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
      Left            =   180
      TabIndex        =   37
      Top             =   7035
      Width           =   7755
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckForUpdate 
         Caption         =   "Check for Update"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCheckClanned_Click()
    Dim blClanCreation As Boolean
    
    If (Len(txtInitiate.text) > 14 Or Len(txtInitiate.text) < 2) Then
        MsgBox "The initiate name must be between 2 to 14 characters.", vbOKOnly Or vbInformation, PROGRAM_TITLE
        Exit Sub
    End If
    
    If (Len(txtChief.text) > 15 Or Len(txtChief.text) < 3) Then
        MsgBox "The chief's name must be from 3 to 15 characters long.", vbOKOnly Or vbInformation, PROGRAM_TITLE
        Exit Sub
    End If
    
    If (Len(txtChiefPass.text) = 0) Then
        MsgBox "Enter a password for the chief.", vbOKOnly Or vbInformation, PROGRAM_TITLE
        Exit Sub
    End If
    
    If (Len(txtInitiatesPassword.text) = 0) Then
        MsgBox "Enter a password for the initiates.", vbOKOnly Or vbInformation, PROGRAM_TITLE
        Exit Sub
    End If
    
    If (Not IsNumeric(txtReconnectTime.text)) Then
        MsgBox "You must enter a numerical value for the reconnect time.", vbOKOnly Or vbInformation, PROGRAM_TITLE
        Exit Sub
    End If
    
    If (Not IsNumeric(txtTimeOut.text)) Then
        MsgBox "You must enter a numerical value for the time out time.", vbOKOnly Or vbInformation, PROGRAM_TITLE
        Exit Sub
    End If
    
    If ((cmbServer.text = vbNullString)) Then
        MsgBox "You must enter a Battle.Net server to connect to.", vbOKOnly Or vbInformation, PROGRAM_TITLE
        Exit Sub
    End If
  
    btnCheckClanned.Enabled = False
    btnManageInitiates.Enabled = False
    chkCustomInitiates.Enabled = False
    btnRefresh.Enabled = True
    btnRefreshKeys.Enabled = False
    cmdResetProxies.Enabled = False
     
    cmbServer.Enabled = False
    txtBNLSServer.Enabled = False
    txtChannel.Enabled = False
    txtTimeOut.Enabled = False
    txtPerProxy.Enabled = False
    txtReconnectTime.Enabled = False
    txtChief.Enabled = False
    txtChiefPass.Enabled = False
    txtClanTag.Enabled = False
    txtClanDescription.Enabled = False
    txtInitiate.Enabled = False
    txtInitiatesPassword.Enabled = False

    With config
        .Channel = txtChannel.text
        .clanTag = txtClanTag.text
        .clanDescription = txtClanDescription.text
        .initiate = txtInitiate.text
        .initiatePassword = txtInitiatesPassword.text
        .reconnectTime = txtReconnectTime.text
        .connectsPerProxy = txtPerProxy.text
        .server = cmbServer.text
    End With
  
    chief.username = txtChief.text
    chief.password = txtChiefPass.text
  
    Dim IP As String, Port As Long, Version As String
    Dim kIdx As Long, arrString() As String
    Dim changeInitiateName As Boolean
    Dim proxy As ProxyData
  
    changeInitiateName = (oldInitiateName = vbNullString _
        Or oldInitiateName <> txtInitiate.text)

    For i = 0 To sckClanMembers.count - 1
        If (bot(i).proxyIP = vbNullString) Then
            proxy = modProxy.getProxy()
          
            With proxy
                IP = .IP
                Port = .Port
                Version = .Version
            End With
            
            If (IP = vbNullString) Then
                AddChat vbRed, "No more proxies available."
                resetAll
                Exit Sub
            End If
    
            bot(i).proxyIP = IP
            bot(i).proxyPort = Port
            bot(i).proxyVersion = Version
        End If

        If (changeInitiateName) Then
            Dim acc As clsAccount
            Set acc = generateInitiate()
      
            bot(i).username = acc.getUsername()
            bot(i).password = acc.getPassword()
        End If
    
        If (bot(i).key = vbNullString) Then
            bot(i).key = keys.getKey(kIdx)
            bot(i).keyIndex = kIdx
      
            If (bot(i).key = vbNullString) Then
                MsgBox "No more keys are available. Unable to create a clan.", vbOKOnly Or vbExclamation, PROGRAM_TITLE
        
                resetAll
                Exit Sub
            End If
        End If
    
        sckClanMembers(i).Connect IP, Port
        tmrInitiateTimeout(i).Enabled = True
  
        AddChat vbYellow, "Initiate #" & i & ": Connecting to " & IP & ":" & Port
    Next i
  
    If (chief.proxyIP = vbNullString) Then
        proxy = modProxy.getProxy()
        
        With proxy
            IP = .IP
            Port = .Port
            Version = .Version
        End With
    
        If (IP = vbNullString) Then
            AddChat vbRed, "No more proxies available."
            resetAll
            Exit Sub
        End If
  
        chief.proxyIP = IP
        chief.proxyPort = Port
        chief.proxyVersion = Version
    End If

    If (chief.key = vbNullString) Then
        chief.key = keys.getKey(kIdx)
        chief.keyIndex = kIdx
      
        If (chief.key = vbNullString) Then
            MsgBox "No more keys are available. Unable to create a clan.", vbOKOnly Or vbExclamation, PROGRAM_TITLE
      
            resetAll
            Exit Sub
        End If
    End If
  
    sckChieftain.Connect IP, Port
    tmrChiefTimeout.Enabled = True

    oldInitiateName = txtInitiate.text
    oldChieftainName = chief.username

    AddChat vbYellow, "Chieftain: Connecting to " & IP & ":" & Port
End Sub

Private Sub btnCheckClanTag_Click()
    Dim blCreateClan As Boolean

    If (Not isValidClanTag(txtClanTag.text)) Then
        MsgBox "Clan tag must be between 2-4 alphanumeric characters.", vbOKOnly Or vbInformation, PROGRAM_TITLE
        Exit Sub
    End If

    If (txtClanDescription.text = vbNullString) Then
        blCreateClan = MsgBox("Really create the clan without a description?", vbYesNo Or vbQuestion, PROGRAM_TITLE)
    
        If (blCreateClan = vbNo) Then Exit Sub
    End If

    config.clanTag = txtClanTag.text
    config.clanDescription = txtClanDescription.text

    btnCheckClanTag.Enabled = False
    txtClanTag.Enabled = False
    txtClanDescription.Enabled = False

    AddChat vbYellow, "Checking clan tag..."
    continueCheckClanTag
End Sub

Private Sub btnCreateClan_Click()
    AddChat vbYellow, "Attempting to create clan..."
    continueCreateClan
  
    btnCreateClan.Enabled = False
End Sub

Private Sub btnManageInitiates_Click()
    frmManageInitiates.Show
End Sub

Private Sub btnRefresh_Click()
    resetAll
    AddChat vbGreen, "All bots have been disconnected."
End Sub

Private Sub btnRefreshKeys_Click()
    loadCDKeys
  
    AddChat vbGreen, "Loaded ", vbYellow, keys.getCount, vbGreen, " CD-Keys!"
End Sub

Private Sub chkCustomInitiates_Click()
    config.useCustomInitiates = chkCustomInitiates.value
End Sub

Private Sub cmdManageInitiates_Click()
    frmManageInitiates.Show
End Sub

Private Sub cmdResetProxies_Click()
    loadProxies
  
    AddChat vbGreen, "Loaded ", vbYellow, modProxy.countProxies, vbGreen, " proxies!"
End Sub

Private Sub Form_Load()
    Me.Caption = Replace(Me.Caption, "%v", PROGRAM_VERSION)
    
    AddChat vbYellow, "Welcome to Simplicity Clan Creator v" & PROGRAM_VERSION & " by Vector"
    
    Hashes(0) = App.path & "\WAR3\Warcraft III.exe"
    
    Dim arrFullList() As Variant
    arrFullList = Array("useast.battle.net", "uswest.battle.net", "europe.battle.net", "asia.battle.net")
    
    Dim idx As Integer: idx = 0
    Dim gateway As Variant

    For Each gateway In arrFullList
        Dim elem As String, realmsString As String, arrIPs() As String
        elem = gateway
    
        arrIPs = Split(Resolve(elem))
      
        cmbServer.AddItem elem
      
        For Each IP In arrIPs
            cmbServer.AddItem IP
        
            If (Not realmsString = vbNullString) Then
                realmsString = realmsString & "|"
            End If
        
            realmsString = realmsString & IP
        Next
    
        dicServerList.Add elem, realmsString
        realmsString = vbNullString
        idx = idx + 1
        If (idx < 4) Then cmbServer.AddItem ""
    Next
  
    If (Dir$(App.path & "\WAR3", vbDirectory) = vbNullString) Then
        MkDir App.path & "\WAR3"
    End If
  
    If (Dir(App.path & "\WAR3\Warcraft III.exe") = vbNullString) Then
        MsgBox "Place Warcraft III.exe in the WAR3 folder, then run this program again.", vbOKOnly Or vbInformation, PROGRAM_TITLE
        unloadAll
        Exit Sub
    End If
  
    If (loadCustomInitiates()) Then
        Dim initiateCount As Integer
        initiateCount = initiateManager.countInitiates()
  
        AddChat vbGreen, "Loaded ", vbYellow, initiateCount, vbGreen, " custom initiate" & IIf(initiateCount > 1, "s", "") & "!"
    End If
  
    loadConfig
  
    If (config.rememberWindowPosition) Then
        If (config.windowTop > 0) Then
            Me.Top = config.windowTop
        End If
    
        If (config.windowLeft > 0) Then
            Me.Left = config.windowLeft
        End If
    End If
  
    If (Dir(App.path & "\CD-Keys.txt") = vbNullString) Then
        MsgBox "No CD-Keys.txt file found. It will be created for you." & vbNewLine _
            & "Place your Warcraft III keys inside of it.", vbOKOnly Or vbInformation, PROGRAM_TITLE
    
        Open App.path & "\CD-Keys.txt" For Output As #1
        Close #1
    Else
        loadCDKeys
    End If
  
    Dim keyCount As Long
    keyCount = keys.getCount()
  
    If (keyCount > 0) Then
        AddChat vbGreen, "Loaded ", vbYellow, keyCount, vbGreen, " CD-Key" & IIf(keyCount > 1, "s", "") & "!"
    End If

    loadProxies
  
    Dim proxyCount As Long
    proxyCount = modProxy.countProxies()
  
    If (proxyCount > 0) Then
        AddChat vbGreen, "Loaded ", vbYellow, proxyCount, vbGreen, " prox" & IIf(proxyCount > 1, "ies", "y") & "!"
    End If
  
    clannedKeyCheckClan = generateClannedKeyCheckClan()
  
    ReDim Packet(8)
    ReDim bot(8)
    Set Chieftain = New clsChieftainData
    Set bnlsPacketHandler = New clsPacketHandler
    
    bnlsPacketHandler.setup sckBNLS, packetType.BNLS
  
    For i = 0 To 8
        If (i > 0) Then
            Load sckClanMembers(i)
            Load tmrInitiateTimeout(i)
            Load tmrReconnect(i)
        End If
    
        Set Packet(i) = New clsPacketHandler
        Packet(i).setup sckClanMembers(i), packetType.BNCS
        tmrInitiateTimeout(i).Interval = config.timeOut
        tmrReconnect(i).Interval = config.reconnectTime
    Next i
    
    tmrChiefTimeout.Interval = config.timeOut
    tmrChiefReconnect.Interval = config.reconnectTime
    Set chiefPacketHandler = New clsPacketHandler
    chiefPacketHandler.setup sckChieftain, packetType.BNCS
  
    If (config.checkUpdateOnStartup) Then
        If (sckUpdateCheck.State = sckClosed) Then
            sckUpdateCheck.Connect "files.codespeak.org", 80
        End If
    End If
  
    programLoaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (programLoaded) Then
        With config
            .reconnectTime = txtReconnectTime.text
            .server = cmbServer.text
            .bnlsServer = txtBNLSServer.text
            .connectsPerProxy = txtPerProxy.text
            .timeOut = txtTimeOut.text
            .Channel = txtChannel.text
            .initiatePassword = txtInitiatesPassword.text
            .initiate = txtInitiate.text
      
            .rememberWindowPosition = chkRememberWindowPosition.value
            .saveClanInfo = chkSaveClanInfo.value
            .checkUpdateOnStartup = chkCheckUpdateOnStartup.value
            .windowTop = frmMain.Top
            .windowLeft = frmMain.Left
        End With
    
        chief.username = txtChief.text
        chief.password = txtChiefPass.text
  
        If (Not config.saveClanInfo) Then
            If (getFileSize(App.path & "\Config.ini") > 0) Then
                Kill App.path & "\Config.ini"
            End If
        End If
  
        saveConfig config.saveClanInfo
        sendBackGoodKeys
    
        unloadAll
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuCheckForUpdate_Click()
    If (sckUpdateCheck.State = sckClosed) Then
        sckUpdateCheck.Connect "files.codespeak.org", 80
        manualUpdateCheck = True
    End If
End Sub

Private Sub mnuQuit_Click()
    unloadAll
End Sub

Private Sub sckBNLS_Connect()
    SEND_BNLS_0x10
End Sub

Private Sub sckBNLS_DataArrival(ByVal bytesTotal As Long)
    Dim data As String, pLen As Long, pID As Byte
    
    sckBNLS.GetData data
    
    Do While Len(data) > 2
        CopyMemory pLen, ByVal Mid$(data, 1, 2), 2
        pID = Asc(Mid$(data, 3, 1))
        bnlsPacketHandler.SetData Mid$(data, 4)
    
        Select Case pID
            Case &H10: RECV_BNLS_0x10
        End Select
    
        data = Mid$(data, pLen + 1)
    Loop
End Sub

Private Sub sckBNLS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AddChat vbRed, "[BNLS] Socket error #" & Number & ": " & Description
    sckBNLS.Close
End Sub

Private Sub sckChieftain_Close()
    chiefError
End Sub

Private Sub sckChieftain_Connect()
    Select Case chief.proxyVersion
        Case "s4"
            sckChieftain.SendData Chr$(&H4) & Chr$(&H1) & Chr$(&H17) & Chr$(&HE0) & stringToChar(returnProperGateway(config.server)) & vbNullString & Chr$(&H0)
        Case "s5"
            sckChieftain.SendData Chr$(&H5) & Chr$(&H1) & Chr$(&H0)
        Case "http"
            sckChieftain.SendData "CONNECT " & returnProperGateway(config.server) & ":6112 HTTP/1.1" & vbCrLf & vbCrLf
    End Select
End Sub

Private Sub sckChieftain_DataArrival(ByVal bytesTotal As Long)
    Dim data As String, pID As Long, pLen As Long

    sckChieftain.GetData data
    If (IsChiefProxyPacket(data)) Then Exit Sub
  
    If (Asc(Left(data, 1)) <> &HFF) Then
        AddChat vbRed, "Chieftain: Unexpected packet received... disconnecting!"
        chiefError
        Exit Sub
    End If
  
    Do While Len(data) > 3
        pID = Asc(Mid(data, 2, 1))
    
        CopyMemory pLen, ByVal Mid$(data, 3, 2), 2
        If (pLen = 0) Then Exit Sub
        chiefPacketHandler.SetData Mid(data, 5, pLen - 4)
  
        Select Case pID
            Case &H0: Chief_Recv0x00
            Case &HA: Chief_Recv0x0A
            Case &H25: Chief_Recv0x25
            Case &H46: Chief_Recv0x46
            Case &H50: Chief_Recv0x50
            Case &H51: Chief_Recv0x51
            Case &H52: Chief_Recv0x52
            Case &H53: Chief_Recv0x53
            Case &H54: Chief_Recv0x54
            Case &H65: Chief_Recv0x65
            Case &H70: Chief_Recv0x70
            Case &H71: Chief_Recv0x71
            Case &H72: Chief_Recv0x72
        End Select
  
        data = Mid(data, pLen + 1)
    Loop
End Sub

Private Sub sckChieftain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    chiefError
End Sub

Private Sub sckClanMembers_Close(index As Integer)
    initiateError index
End Sub

Private Sub sckClanMembers_Connect(index As Integer)
    Select Case bot(index).proxyVersion
        Case "s4"
            sckClanMembers(index).SendData Chr$(&H4) & Chr$(&H1) & Chr$(&H17) & Chr$(&HE0) & stringToChar(returnProperGateway(config.server)) & vbNullString & Chr$(&H0)
        Case "s5"
            sckClanMembers(index).SendData Chr$(&H5) & Chr$(&H1) & Chr$(&H0)
        Case "http"
            sckClanMembers(index).SendData "CONNECT " & returnProperGateway(config.server) & ":6112 HTTP/1.1" & vbCrLf & vbCrLf
    End Select
End Sub

Private Sub sckClanMembers_DataArrival(index As Integer, ByVal bytesTotal As Long)
    Dim data As String, pID As Long, pLen As Long
  
    sckClanMembers(index).GetData data
    If (IsProxyPacket(index, data)) Then Exit Sub
  
    If (Asc(Left(data, 1)) <> &HFF) Then
        AddChat vbRed, "Socket #" & index & ": Unexpected packet received... disconnecting!"
        initiateError index
        Exit Sub
    End If
  
    Do While Len(data) > 3
        pID = Asc(Mid(data, 2, 1))
    
        CopyMemory pLen, ByVal Mid$(data, 3, 2), 2
        If (pLen = 0) Then Exit Sub
        Packet(index).SetData Mid(data, 5, pLen - 4)
    
        Select Case pID
            Case &H0: Recv0x00 index
            Case &HA: Recv0x0A index
            Case &H25: Recv0x25 index
            Case &H46: Recv0x46 index
            Case &H50: Recv0x50 index
            Case &H51: Recv0x51 index
            Case &H52: Recv0x52 index
            Case &H53: Recv0x53 index
            Case &H54: Recv0x54 index
            Case &H65: Recv0x65 index
            Case &H70: Recv0x70 index
            Case &H72: Recv0x72 index
        End Select
    
        data = Mid(data, pLen + 1)
    Loop
End Sub

Private Sub sckClanMembers_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    initiateError index
End Sub

Private Sub sckUpdateCheck_Connect()
    sckUpdateCheck.SendData "GET /projects/simplicity/CurrentVersion.txt HTTP/1.1" & vbCrLf _
                            & "User-Agent: Simplicity/" & PROGRAM_VERSION & vbCrLf _
                            & "Host: files.codespeak.org" & vbCrLf & vbCrLf
End Sub

Private Sub sckUpdateCheck_DataArrival(ByVal bytesTotal As Long)
    Dim data As String, ver As String
    sckUpdateCheck.GetData data

    updateString = updateString & data

    tmrCheckUpdate.Enabled = True
End Sub

Private Sub sckUpdateCheck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AddChat vbRed, "Unable to check for update!"
    sckUpdateCheck.Close
End Sub

Private Sub tmrChiefReconnect_Timer()
    sckChieftain.Connect chief.proxyIP, chief.proxyPort
    tmrChiefReconnect.Enabled = False
End Sub

Private Sub tmrChiefTimeout_Timer()
    chiefError
End Sub

Private Sub tmrQueue_Timer()
    If (chiefData.getFriendsCount() = 25) Then
        tmrQueue.Enabled = False
      
        MsgBox "Chief's friends list is full. You need to remove some friends to continue.", vbOKOnly Or vbInformation, PROGRAM_TITLE
        
        Dim friendsExcludeInitiates As Integer
        friendsExcludeInitiates = chiefData.getFriendsWithoutInitiatesCount()
      
        frmFriendsList.addFriends (friendsExcludeInitiates + 9) - 25
        frmFriendsList.Show
          
        Exit Sub
    Else
        Dim initiate As String
        initiate = chiefData.popQueue()
      
        With chiefPacketHandler
            .InsertNTString "/friends add " & initiate
            .sendPacket &HE
        End With
        chiefData.addFriend initiate
          
        Dim botIndex As Integer
        botIndex = modOtherCode.getBotIndexByName(initiate)
    
        AddChat vbYellow, "Added " & IIf(bot(botIndex).hasRestrictedKey, "restricted ", "") & "initiate """ & initiate & """ to chief's friends list."
    End If
      
    If (chiefData.isQueueEmpty()) Then
        If (chiefData.isReplacingFriends()) Then
            AddChat vbGreen, "Finished replacing friends on the chieftain's friends list."
            AddChat vbGreen, "You may now check the clan tag."
        Else
            AddChat vbGreen, "All bots are ready. You may now check the clan tag."
        End If
      
        btnCheckClanTag.Enabled = True
        tmrQueue.Enabled = False
    End If
End Sub

Private Sub tmrReconnect_Timer(index As Integer)
    sckClanMembers(index).Connect bot(index).proxyIP, bot(index).proxyPort
    tmrReconnect(index).Enabled = False
End Sub

Private Sub tmrInitiateTimeout_Timer(index As Integer)
    initiateError index
End Sub

Private Sub tmrCheckUpdate_Timer()
    On Error GoTo err
  
    Dim versionToCheck As String, updateMsg As String, msgBoxResult As Integer
  
    versionToCheck = Split(updateString, "Content-Type: text/plain" & vbCrLf & vbCrLf)(1)

    If (isNewVersion(versionToCheck)) Then
        updateMsg = "There is a new update for Simplicity!" & vbNewLine & vbNewLine & "Your version: " & PROGRAM_VERSION & " new version: " & versionToCheck & vbNewLine & vbNewLine _
                & "Would you like to visit the downloads page for updates?"

        msgBoxResult = MsgBox(updateMsg, vbYesNo Or vbInformation, "New Simplicity version available!")

        If (msgBoxResult = vbYes) Then
            ShellExecute 0, "open", RELEASES_URL, vbNullString, vbNullString, 4
        End If
    Else
        If (manualUpdateCheck) Then
            MsgBox "There is no new version at this time.", vbOKOnly Or vbInformation, PROGRAM_TITLE
            manualUpdateCheck = False
        End If
    End If
err:
    If (err.Number > 0) Then
        err.Clear
        AddChat vbRed, "Unable to check for update!"
    End If

    updateString = vbNullString
    tmrCheckUpdate.Enabled = False
    sckUpdateCheck.Close
End Sub

Public Sub initiateError(ByVal index As Integer)
    Dim IP As String, Port As Long
    Dim Version As String
    
    sckClanMembers(index).Close
    tmrInitiateTimeout(index).Enabled = False
    
    AddChat vbRed, "Initiate #" & index & ": Connection to proxy failed."
      
    If (bot(index).isReadyForPreparation) Then
        connectedCount = connectedCount - 1
        lblConnected.Caption = "Connected: " & Right(" " & connectedCount, 2)
    End If
  
    bot(index).isReadyForPreparation = False
    btnCreateClan.Enabled = False
    btnCheckClanTag.Enabled = False
    bot(index).loggedOn = False
    bot(index).hasCheckedFriendsList = False
      
    bot(index).proxyIP = vbNullString
    bot(index).proxyPort = 0
    bot(index).proxyVersion = vbNullString
    
    Do
        Dim proxy As ProxyData
        proxy = modProxy.getProxy()
  
        With proxy
            IP = .IP
            Port = .Port
            Version = .Version
        End With
    
        If (IP = vbNullString) Then
            AddChat vbRed, "No more proxies available."
            resetAll
            Exit Sub
        End If
    Loop While IP = bot(index).proxyIP
    
    With bot(index)
        .proxyIP = IP
        .proxyPort = Port
        .proxyVersion = Version
    End With
  
    sckClanMembers(index).Connect IP, Port
    tmrInitiateTimeout(index).Enabled = True
    AddChat vbYellow, "Initiate #" & index & ": Connecting to " & IP & ":" & Port
End Sub

Public Sub chiefError()
    Dim IP As String, Port As Long
    Dim Version As String
    
    AddChat vbRed, "Chieftain: Connection to proxy failed."
    tmrChiefTimeout.Enabled = False
    sckChieftain.Close
      
    If (chiefData.isQueueEnabled) Then
        tmrQueue.Enabled = False
    End If
      
    If (chief.isReadyForPreparation) Then
        connectedCount = connectedCount - 1
        lblConnected.Caption = "Connected: " & Right(" " & connectedCount, 2)
    End If
    
    chief.isReadyForPreparation = False
    btnCreateClan.Enabled = False
    btnCheckClanTag.Enabled = False
    chief.loggedOn = False
    chiefData.clearFriends
    
    chief.proxyIP = vbNullString
    chief.proxyPort = 0
    chief.proxyVersion = vbNullString
  
    Do
        Dim proxy As ProxyData
        proxy = modProxy.getProxy()
    
        With proxy
            IP = .IP
            Port = .Port
            Version = .Version
        End With
      
        If (IP = vbNullString) Then
            AddChat vbRed, "No more proxies available."
            resetAll
            Exit Sub
        End If
    Loop While IP = chief.proxyIP
      
    With chief
        .proxyIP = IP
        .proxyPort = Port
        .proxyVersion = Version
    End With
    
    tmrChiefTimeout.Enabled = True
    AddChat vbYellow, "Chieftain: Connecting to " & IP & ":" & Port
    sckChieftain.Connect IP, Port
End Sub

Public Sub continueCreateClan()
    Dim initCount As Integer
    initCount = dicInitiatesAdded.count

    With chiefPacketHandler
        .InsertDWORD &H0
        .InsertNTString config.clanDescription
        .InsertNonNTString StrReverse$(Left$(config.clanTag & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0), 4))
        .InsertByte initCount
    
        For Each key In dicInitiatesAdded.keys
            .InsertNTString key
        Next
    
        .sendPacket &H71
    End With
End Sub

Public Sub continueCheckClanTag()
    isCheckingClanTag = True
  
    With chiefPacketHandler
        .InsertDWORD &H0
        .InsertNonNTString StrReverse$(Left$(config.clanTag & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0), 4))
        .sendPacket &H70
    End With
End Sub


