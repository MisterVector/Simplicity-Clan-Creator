Attribute VB_Name = "modOtherCode"
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function GetCurrentProcess Lib "Kernel32" () As Long
Public Declare Function EmptyWorkingSet Lib "psapi.dll" (ByVal hProcess As Long) As Long
Public Declare Function SetProcessWorkingSetSize Lib "Kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long _
                                                                , ByVal dwMaximumWorkingSetSize As Long) As Long

Public Function ews_memory() As Long:   ews_memory = EmptyWorkingSet(GetCurrentProcess):                  End Function
Public Function spw_memory() As Long:   spw_memory = SetProcessWorkingSetSize(GetCurrentProcess, -1, -1): End Function
Public Sub FreeMemory()
  ews_memory
  spw_memory
End Sub

Public Function stringToChar(ByVal str As String) As String
  Dim arr() As String
  
  arr = Split(str, ".")
  For i = 0 To UBound(arr)
    stringToChar = stringToChar & Chr$(CStr(arr(i)))
  Next i
End Function

Public Sub resetAll(Optional ByVal supressMessage As Boolean = False)
  Set Chieftain = New clsChieftainData
  isCheckingClanTag = False
  initiateNumber = 1
  initiateManager.resetIndex
  botsAdded.RemoveAll
  dicInitiatesAdded.RemoveAll
  
  frmMain.tmrQueue.Enabled = False
  For i = 0 To 8
    bot(i).hasDoneProcedure = False
    bot(i).loggedOn = False
    bot(i).isReadyForPreparation = False
    frmMain.sckClanMembers(i).Close
    frmMain.tmrReconnect(i).Enabled = False
    frmMain.tmrInitiateTimeout(i).Enabled = False
  Next i
  frmMain.sckChieftain.Close
  
  frmMain.sckBNLS.Close
  
  frmMain.tmrChiefReconnect.Enabled = False
  frmMain.tmrChiefTimeout.Enabled = False
  
  frmMain.btnCheckClanTag.Enabled = False
  frmMain.btnCreateClan.Enabled = False
  frmMain.btnCheckClanned.Enabled = True
  frmMain.btnManageInitiates.Enabled = True
  frmMain.btnRefresh.Enabled = False
  frmMain.btnRefreshKeys.Enabled = True
  frmMain.cmdResetProxies.Enabled = True
  
  frmMain.cmbServer.Enabled = True
  frmMain.txtBNLSServer.Enabled = True
  frmMain.txtChannel.Enabled = True
  frmMain.txtTimeOut.Enabled = True
  frmMain.txtPerProxy.Enabled = True
  frmMain.txtReconnectTime.Enabled = True
  frmMain.txtChief.Enabled = True
  frmMain.txtChiefPass.Enabled = True
  frmMain.txtClanTag.Enabled = True
  frmMain.txtClanDescription.Enabled = True
  frmMain.txtInitiate.Enabled = True
  frmMain.txtInitiatesPassword.Enabled = True
  
  frmMain.chkCustomInitiates.Enabled = True

  chief.loggedOn = False
  chief.hasDoneProcedure = False
  chief.isReadyForPreparation = False
  chiefData.clearFriends
  
  connectedCount = 0
  frmMain.lblConnected.Caption = "Connected:  0"
  
  FreeMemory
  
  If (Not supressMessage) Then
    AddChat vbGreen, "Simplicity has been refreshed."
  End If
End Sub

Public Function generateClannedKeyCheckClan() As String
  Dim alphanumeric As String, final As String, idx As Integer
  alphanumeric = "abcdefghijklmnopqrstuvwxyz1234567890"
  
  For i = 0 To 3
    idx = Int(Rnd(GetTickCount()) * Len(alphanumeric))
    If idx = 0 Then idx = 1
  
    final = final & Mid(alphanumeric, idx, 1)
  Next i
  
  generateClannedKeyCheckClan = final
End Function

Public Sub prepareInitiatesAndChief()
  Dim mustFriendChief As Boolean, mustFriendInitiate As Boolean

  AddChat vbYellow, "Preparing chief and initiates for tag check..."

  Dim i As Integer

  For i = 0 To 8
    If ((bot(i).hasRestrictedKey Or chief.hasRestrictedKey) _
        And Not bot(i).hasChieftainAsFriend) Then
      With Packet(i)
        .InsertNTString "/friends add " & chief.username
        .sendPacket i, &HE
      End With
    End If

    If ((bot(i).hasRestrictedKey Or chief.hasRestrictedKey) _
        And Not chiefData.isFriend(bot(i).username)) Then
      chiefData.addQueue bot(i).username
    End If
  Next i
  
  If Not chiefData.isQueueEmpty() Then
    frmMain.tmrQueue.Enabled = True
  Else
    AddChat vbYellow, "All bots are ready. You may now check the clan tag."
    frmMain.btnCheckClanTag.Enabled = True
  End If
End Sub

Public Function countReadyForPreparation(Optional ByVal includeChief As Boolean = True) As Integer
  Dim count As Integer

  For i = 0 To 8
    If bot(i).isReadyForPreparation Then
      count = count + 1
    End If
  Next i
  
  If includeChief And chief.isReadyForPreparation Then
    count = count + 1
  End If
  
  countReadyForPreparation = count
End Function

Public Function getBotIndexByName(ByVal user As String) As Integer
  For i = 0 To 8
    If LCase(bot(i).username) = LCase(user) Then
      getBotIndexByName = i
      Exit Function
    End If
  Next i
End Function

Public Sub AddChat(ParamArray saElements() As Variant)
  With frmMain.rtbChat
    Dim sp() As String
    sp = Split(.text, vbNewLine)
    
    If UBound(sp) >= 49 Then
      .text = vbNullString
    End If
    
    .SelStart = Len(.text)
    .SelLength = 0
    .SelColor = vbWhite
    .SelText = "[" & Time() & "] "
    
    For i = 0 To UBound(saElements) Step 2
      .SelStart = Len(.text)
      .SelLength = 0
      .SelColor = saElements(i)
      .SelText = saElements(i + 1) & IIf(i + 1 = UBound(saElements), vbNewLine, "")
    Next i
  End With
End Sub

Public Function returnProperGateway(ByVal gateway As String) As String
  Dim gatewayList() As String
  
  If IsNumeric(Replace(gateway, ".", "")) Then
    returnProperGateway = gateway
    Exit Function
  End If

  gatewayList = Split(Resolve(gateway))
  returnProperGateway = gatewayList(CInt(Rnd * UBound(gatewayList)))
End Function

Public Function getFileSize(path As String) As Long
  On Error GoTo oops:
  
  Open path For Input As #1
  getFileSize = CLng(LOF(1))
  Close #1
  
oops:
  If getFileSize > 0 Then Exit Function
  getFileSize = 0
End Function

Public Sub disconnectInitiate(ByVal index As Integer)
  frmMain.sckClanMembers(index).Close
  frmMain.tmrInitiateTimeout(index).Enabled = False
  
  If bot(index).isReadyForPreparation Then
    connectedCount = connectedCount - 1
    frmMain.lblConnected.Caption = "Connected: " & Right(" " & connectedCount, 2)
  End If
  
  bot(index).isReadyForPreparation = False
  frmMain.btnCreateClan.Enabled = False
  frmMain.btnCheckClanTag.Enabled = False
  bot(index).loggedOn = False
  bot(index).hasCheckedFriendsList = False
End Sub

Public Sub reconnectInitiate(ByVal index As Integer)
  Dim IP As String, Port As Long
  Dim version As String
  
  Do
    Dim proxy As ProxyData
    proxy = modProxy.getProxy()
  
    With proxy
      IP = .IP
      Port = .Port
      version = .version
    End With
    
    If IP = vbNullString Then
      AddChat vbRed, "No more proxies available."
      resetAll
      Exit Sub
    End If
  Loop While IP = bot(index).proxyIP
    
  With bot(index)
    .proxyIP = IP
    .proxyPort = Port
    .proxyVersion = version
  End With
  
  frmMain.sckClanMembers(index).Connect IP, Port
  frmMain.tmrInitiateTimeout(index).Enabled = True
  AddChat vbYellow, "Initiate #" & index & ": Connecting to " & IP & ":" & Port
End Sub

Public Function IsProxyPacket(index As Integer, ByVal data As String) As Boolean
  Select Case Mid$(data, 1, 2)
    Case Chr(&H0) & Chr(&H5A): 'Accepted
                               frmMain.sckClanMembers(index).SendData Chr$(&H1)
                               Send0x50 index
                               IsProxyPacket = True
                               Exit Function
    Case Chr(&H0) & Chr(&H5B): 'Denied
                               IsProxyPacket = True
    Case Chr(&H0) & Chr(&H5C): 'Rejected
                               IsProxyPacket = True
    Case Chr(&H0) & Chr(&H5D): 'Rejected
                               IsProxyPacket = True
  End Select

  If Not IsProxyPacket Then
    If Len(data) >= 12 And LCase(Left(data, 4)) = "http" Then
      Dim packetOutput As String
      Dim responseCode As String
      
      responseCode = Mid(data, 10, 3)
    
      Select Case responseCode
        Case "200"
          frmMain.sckClanMembers(index).SendData Chr$(&H1)
          Send0x50 index
          IsProxyPacket = True
        Case "307"
          frmMain.initiateError index
          IsProxyPacket = True
      End Select
    End If
  End If
End Function

Public Function IsChiefProxyPacket(ByVal data As String) As Boolean
  Select Case Mid$(data, 1, 2)
    Case Chr(&H0) & Chr(&H5A): 'Accepted
                               frmMain.sckChieftain.SendData Chr$(&H1)
                               Chief_Send0x50
                               IsChiefProxyPacket = True
                               Exit Function
    Case Chr(&H0) & Chr(&H5B): 'Denied
                               IsChiefProxyPacket = True
    Case Chr(&H0) & Chr(&H5C): 'Rejected
                               IsChiefProxyPacket = True
    Case Chr(&H0) & Chr(&H5D): 'Rejected
                               IsChiefProxyPacket = True
  End Select

  If Not IsChiefProxyPacket Then
    If Len(data) >= 12 And LCase(Left(data, 4)) = "http" Then
      Dim packetOutput As String
      Dim responseCode As String
      
      responseCode = Mid(data, 10, 3)
    
      Select Case responseCode
        Case "200"
          frmMain.sckChieftain.SendData Chr$(&H1)
          Chief_Send0x50
          IsChiefProxyPacket = True
        Case "307"
          frmMain.chiefError
          IsChiefProxyPacket = True
      End Select
    End If
  End If
End Function

Public Function isInitiate(name As String) As Boolean
  For i = 0 To 8
    If LCase(bot(i).username) = LCase(name) Then
      isInitiate = True
      Exit Function
    End If
  Next i
  
  isInitiate = False
End Function

Public Function getRealmName(ByVal serverConfig As String) As String
  Dim tempServer As String, IPs() As String, doExit As Boolean
  
  For Each key In dicServerList.keys
    If LCase(key) = LCase(serverConfig) Then
      tempServer = key
      Exit For
    End If
    
    IPs = Split(dicServerList.Item(key), "|")
    
    For Each IP In IPs
      If IP = serverConfig Then
        tempServer = key
        doExit = True
        Exit For
      End If
    Next
    
    If doExit Then
      Exit For
    End If
  Next

  Select Case tempServer
    Case "uswest.battle.net": getRealmName = "Lordaeron"
    Case "useast.battle.net": getRealmName = "Azaroth"
    Case "europe.battle.net": getRealmName = "Northrend"
    Case "asia.battle.net": getRealmName = "Kalimdor"
  End Select
End Function

Public Function KillNull(ByVal text As String) As String
  Dim findNull As Integer
  
  findNull = InStr(text, Chr(0))
  KillNull = IIf(findNull > 0, Mid(text, 1, findNull - 1), text)
End Function

Public Sub unloadAll()
  Dim oFrm As Form
  
  For Each oFrm In Forms
    Unload oFrm
  Next
End Sub

Public Function isValidClanTag(clanTag As String) As Boolean
  Dim lenTag, ch, ascii
  
  lenTag = Len(clanTag)
  
  If lenTag > 4 Or lenTag < 2 Then
    isValidClanTag = False
    Exit Function
  End If

  For i = 1 To lenTag
    ch = Mid(clanTag, i, 1)
    ascii = Asc(ch)
    
    If (Not IsNumeric(ch) And Not ((ascii >= 65 And ascii <= 90) Or (ascii >= 97 And ascii <= 122))) Then
      isValidClanTag = False
      Exit Function
    End If
  Next i
  
  isValidClanTag = True
End Function

Public Function generateInitiate() As clsAccount
  Dim acc As New clsAccount, initFound As Boolean

  ' initial initialization
  If initiateNumber = 0 Then
    initiateNumber = 1
  End If
  
  If config.useCustomInitiates Then
    Dim tempInitiate As clsAccount
    Set tempInitiate = initiateManager.getNextInitiate()
    
    If Not tempInitiate Is Nothing Then
      acc.setUsername (tempInitiate.getUsername())
      acc.setPassword (tempInitiate.getPassword())
    
      initFound = True
    End If
  End If
  
  If Not initFound Then
    acc.setUsername (config.initiate & initiateNumber)
    acc.setPassword (config.initiatePassword)
    
    initiateNumber = initiateNumber + 1
  End If
  
  Set generateInitiate = acc
End Function

Public Sub exportClanToFile()
  Open App.path & "\CreatedClans.txt" For Append As #1
    Print #1, "Clan " & config.clanTag & " @ " & getRealmName(config.server) & " Chieftain: " & chief.username & " Password: " _
            & chief.password & " Initiate password: " & config.initiatePassword & " " _
            & "Created on " & Date & " at " & Time() & "."
  Close #1
End Sub
