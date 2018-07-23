Attribute VB_Name = "modClanFunctions"
Public Function generateClannedKeyCheckClan() As String
    Dim alphanumeric As String, final As String, idx As Integer
    alphanumeric = "abcdefghijklmnopqrstuvwxyz1234567890"
  
    For i = 0 To 3
        idx = Int(Rnd(GetTickCount()) * Len(alphanumeric))
        If (idx = 0) Then idx = 1
  
        final = final & Mid$(alphanumeric, idx, 1)
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
                .sendPacket &HE
            End With
        End If

        If ((bot(i).hasRestrictedKey Or chief.hasRestrictedKey) _
                And Not chiefData.isFriend(bot(i).username)) Then
            chiefData.addQueue bot(i).username
        End If
    Next i
  
    If (Not chiefData.isQueueEmpty()) Then
        frmMain.tmrQueue.Enabled = True
    Else
        AddChat vbYellow, "All bots are ready. You may now check the clan tag."
        frmMain.btnCheckClanTag.Enabled = True
    End If
End Sub

Public Function countReadyForPreparation(Optional ByVal includeChief As Boolean = True) As Integer
    Dim count As Integer

    For i = 0 To 8
        If (bot(i).isReadyForPreparation) Then
            count = count + 1
        End If
    Next i
  
    If (includeChief And chief.isReadyForPreparation) Then
        count = count + 1
    End If
  
    countReadyForPreparation = count
End Function

Public Sub disconnectInitiate(ByVal index As Integer)
    frmMain.sckClanMembers(index).Close
    frmMain.tmrInitiateTimeout(index).Enabled = False
  
    If (bot(index).isReadyForPreparation) Then
        connectedCount = connectedCount - 1
        frmMain.lblConnected.Caption = "Connected: " & Right$(" " & connectedCount, 2)
    End If
  
    bot(index).isReadyForPreparation = False
    frmMain.btnCreateClan.Enabled = False
    frmMain.btnCheckClanTag.Enabled = False
    bot(index).loggedOn = False
    bot(index).hasCheckedFriendsList = False
End Sub

Public Sub reconnectInitiate(ByVal index As Integer)
    Dim IP As String, Port As Long
    Dim Version As String
  
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
    
    frmMain.sckClanMembers(index).Connect IP, Port
    frmMain.tmrInitiateTimeout(index).Enabled = True
    AddChat vbYellow, "Initiate #" & index & ": Connecting to " & IP & ":" & Port
End Sub

Public Function isInitiate(name As String) As Boolean
    For i = 0 To 8
        If (LCase$(bot(i).username) = LCase$(name)) Then
            isInitiate = True
            Exit Function
        End If
    Next i
  
    isInitiate = False
End Function

Public Function isValidClanTag(clanTag As String) As Boolean
    Dim lenTag, ch, ascii
  
    lenTag = Len(clanTag)
  
    If (lenTag > 4 Or lenTag < 2) Then
        isValidClanTag = False
        Exit Function
    End If

    For i = 1 To lenTag
        ch = Mid$(clanTag, i, 1)
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
    If (initiateNumber = 0) Then
        initiateNumber = 1
    End If
  
    If (config.useCustomInitiates) Then
        Dim tempInitiate As clsAccount
        Set tempInitiate = initiateManager.getNextInitiate()
    
        If (Not tempInitiate Is Nothing) Then
            acc.setUsername (tempInitiate.getUsername())
            acc.setPassword (tempInitiate.getPassword())
    
            initFound = True
        End If
    End If
  
    If (Not initFound) Then
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
