Attribute VB_Name = "modBNET"
Public Sub Send0x00(index As Integer)
    Packet(index).sendPacket &H0
End Sub

Public Sub Recv0x00(index As Integer)
    Send0x00 index
End Sub

Public Sub Recv0x25(index As Integer)
    Send0x25 index
End Sub

Public Sub Send0x25(index As Integer)
    With Packet(index)
        .InsertDWORD .GetDWORD
        .sendPacket &H25
    End With
End Sub

Public Sub Send0x50(index As Integer)
    With Packet(index)
        .InsertDWORD &H0
        .InsertNonNTString "68XI3RAW"
        .InsertDWORD config.verByte
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertNTString "USA"
        .InsertNTString "United States"
        .sendPacket &H50
    End With
End Sub

Public Sub Recv0x50(index As Integer)
    Dim clientToken As Long, serverToken As Long
    Dim MPQFileN As String, valueString As String
    Dim FT As FILETIME, archiveTime As String

    Packet(index).Skip 4              'Logon type
  
    clientToken = GetTickCount
    serverToken = Packet(index).GetDWORD
    Packet(index).Skip 4    'UDPValue
    
    FT.dwLowDateTime = Packet(index).GetDWORD
    FT.dwHighDateTime = Packet(index).GetDWORD
    archiveTime = GetFTTime(FT)
    
    MPQFileN = Packet(index).getNTString
    valueString = Packet(index).getNTString
    
    Send0x51 index, clientToken, serverToken, MPQFileN, valueString, archiveTime
End Sub

Public Sub Send0x51(index As Integer, clientToken As Long, serverToken As Long, MPQFileN As String, valueString As String, archiveTime As String)
    Dim EXEInfo      As String: EXEInfo = String$(crev_max_result, Chr$(0))
    Dim Checksum     As Long
    Dim Version      As Long
    Dim CDKeyHash As String * 20
    Dim ProdVal   As Long
    Dim PubVal    As Long
    
    If (decode_hash_cdkey(bot(index).key, clientToken, serverToken, PubVal, ProdVal, CDKeyHash) = 0) Then
        MsgBox "Unable to decode your CD-Keys.", vbOKOnly Or vbCritical, PROGRAM_TITLE
        Unload All

        Exit Sub
    End If

    check_revision archiveTime, MPQFileN, valueString, App.path & "\VersionCheck.ini", "WAR3", Version, Checksum, EXEInfo

    With Packet(index)
        .InsertDWORD clientToken
        .InsertDWORD Version
        .InsertDWORD Checksum
        .InsertDWORD &H1
        .InsertDWORD &H0
        
        .InsertDWORD Len(bot(index).key)
        .InsertDWORD ProdVal
        .InsertDWORD PubVal
        .InsertDWORD &H0
        .InsertNonNTString CDKeyHash
        
        .InsertNTString KillNull(EXEInfo)
        .InsertNTString "Simplicity"
        .sendPacket &H51
    End With
End Sub

Public Sub Recv0x51(index As Integer)
    Dim getStatusCode As Long, keyIdx As Long
  
    frmMain.tmrInitiateTimeout(index).Enabled = False
    getStatusCode = Packet(index).GetDWORD
  
    If (getStatusCode = &H0) Then
        Send0x53 index
    Else
        frmMain.sckClanMembers(index).Close
  
        Select Case getStatusCode
            Case &H100, &H101
                resetAll True
        
                If (attemptedVerByteUpdate) Then
                    AddChat vbRed, "Warcraft III hashes are out of date. You need to update them."
                Else
                    AddChat vbRed, "Initiate #" & index & ": Your game version is out of date."
                    AddChat vbYellow, "Attempting to update version byte..."
          
                    If (config.bnlsServer <> vbNullString) Then
                        frmMain.sckBNLS.Connect config.bnlsServer, 9367
                    Else
                        AddChat vbRed, "BNLS server not configured. Unable to update version byte!"
                    End If
                End If
        
                Exit Sub
            Case &H102
                AddChat vbRed, "Initiate #" & index & ": Your game version must be downgraded."
                resetAll
                Exit Sub
            Case &H201
                AddChat vbRed, "Initiate #" & index & ": Your CD-Key is in use by " & Packet(index).getNTString & "."
                addKey bot(index).keyIndex, bot(index).key, KeyType.IN_USE
            Case &H202, &H203
                If (getStatusCode = &H202) Then
                    AddChat vbRed, "Initiate #" & index & ": That CD-Key is banned."
                Else
                    AddChat vbRed, "Initiate #" & index & ": That CD-Key has the wrong product."
                End If
        
                addKey bot(index).keyIndex, bot(index).key, KeyType.BAD
        End Select
    
        bot(index).key = keys.getKey(keyIdx)
        bot(index).keyIndex = keyIdx
    
        If (bot(index).key = vbNullString) Then
            AddChat vbRed, "Initiate #" & index & ": You have run out of keys."
            resetAll
        Else
            frmMain.tmrReconnect(index).Enabled = True
        End If
    End If
End Sub

Public Sub Send0x53(index As Integer)
    Dim nls_A As String

    bot(index).nls_P = nls_init(bot(index).username, bot(index).password)

    If (bot(index).nls_P = 0) Then
        MsgBox "NLS made a bad call.", vbOKOnly Or vbCritical, PROGRAM_TITLE
        unloadAll
        Exit Sub
    End If

    nls_A = Space$(Len(bot(index).username) + 33)
    
    If (nls_account_logon(bot(index).nls_P, nls_A) = 0) Then
        MsgBox "Unable to create NLS key.", vbOKOnly Or vbCritical, PROGRAM_TITLE
    
        unloadAll
        Exit Sub
    End If

    Packet(index).InsertNonNTString Left$(nls_A, Len(nls_A) - Len(bot(index).username) - 1)
    Packet(index).InsertNTString bot(index).username
    Packet(index).sendPacket &H53
End Sub

Public Sub Recv0x53(index As Integer)
    Select Case Packet(index).GetDWORD
        Case &H0: Send0x54 index   'Passed
        Case &H1: Send0x52 index   'Account Not made
        Case Else
            AddChat vbYellow, "Initiate #" & index & ": Account error..."
            resetAll
    End Select
End Sub

Public Sub Send0x52(index As Integer)
    Dim SaltHash As String: SaltHash = Space$(Len(bot(index).username) + 65)

    If (nls_account_create(bot(index).nls_P, SaltHash) = 0) Then
        MsgBox "Unable to create NLS salt.", vbOKOnly Or vbCritical, PROGRAM_TITLE  'Unable to create salt
        Exit Sub
    End If

    Packet(index).InsertNonNTString SaltHash
    Packet(index).sendPacket &H52
End Sub

Public Sub Recv0x52(index As Integer)
    Dim result As Long
  
    result = Packet(index).GetDWORD
  
    If (result = &H0) Then
        Send0x53 index
        AddChat vbGreen, "Initiate #" & index & ": Account successfully created."
    Else
        frmMain.sckClanMembers(index).Close
        AddChat vbRed, "Socket #" & index & ": Account creation error."
        frmMain.tmrReconnect(index).Enabled = True
    End If
End Sub

Public Sub Send0x54(index As Integer)
    Dim ProofHash As String * 20
    Dim Salt      As String: Salt = Packet(index).GetNonNTString(32)
    Dim ServerKey As String: ServerKey = Packet(index).GetNonNTString(32)

    nls_account_logon_proof bot(index).nls_P, ProofHash, ServerKey, Salt

    Packet(index).InsertNonNTString ProofHash
    Packet(index).sendPacket &H54
End Sub

Public Sub Recv0x54(index As Integer)
    Dim errorCode As Long, arrString() As String
    errorCode = Packet(index).GetDWORD
  
    Select Case errorCode
        Case &H0: GoTo Continue
        Case &H1:
        Case &H2:
            AddChat vbRed, "Initiate #" & index & ": Invalid password for " & bot(index).username & "."
      
            Dim acc As New clsAccount
            Set acc = generateInitiate()
      
            bot(index).username = acc.getUsername()
            bot(index).password = acc.getPassword()
      
            AddChat vbYellow, "Initiate #" & index & ": Attempting to use name " & bot(index).username
      
            frmMain.sckClanMembers(index).Close
            frmMain.tmrReconnect(index).Enabled = True
            Exit Sub
        Case &HF:
        Case &HE:
            GoTo Continue
    End Select
    
    resetAll
    Exit Sub
Continue:
    nls_free (bot(index).nls_P)        'Unloads the NLS object to avoid overhead

    Send0x0A index
End Sub

Public Sub Send0x65(index As Integer)
    With Packet(index)
        .sendPacket &H65
    End With
End Sub

Public Sub Recv0x65(index As Integer)
    Dim friendsCount As Integer, curFriend As String
  
    With Packet(index)
        friendsCount = .GetByte
    
        If (friendsCount > 0) Then
            For i = 0 To friendsCount - 1
                curFriend = .getNTString
        
                If (LCase$(curFriend) = LCase$(chief.username)) Then
                    bot(index).hasChieftainAsFriend = True
                End If
        
                .Skip 6
                .getNTString
            Next i
        End If
    End With
  
    bot(index).hasCheckedFriendsList = True
  
    Send0x70 index
End Sub

Public Sub Send0x70(index As Integer)
    Dim sendTag As String
  
    For i = 0 To 3
        sendTag = sendTag & Chr$(Asc(Mid$(clannedKeyCheckClan, i + 1, 1)))
    Next i
 
    With Packet(index)
        .InsertDWORD &H0
        .InsertNonNTString StrReverse$(sendTag)
        .sendPacket &H70
    End With
End Sub

Public Sub Recv0x70(index As Integer)
    Dim result As Long, count As Integer, keyIdx As Long
    Dim getInitiate As String, getRemainingCount As Integer
  
    Packet(index).GetDWORD
    result = Packet(index).GetByte
  
    Select Case result
        Case &H0
            AddChat vbGreen, "Initiate #" & index & ": Found a non-clanned " & IIf(bot(index).hasRestrictedKey, "restricted ", vbNullString) & "key!"
            bot(index).isReadyForPreparation = True
      
            connectedCount = connectedCount + 1
            frmMain.lblConnected.Caption = "Connected: " & Right$(" " & connectedCount, 2)
      
            Dim readyCount As Integer
            readyCount = countReadyForPreparation()
      
            If (readyCount = 10) Then
                prepareInitiatesAndChief
            End If
        Case &H1
            clannedKeyCheckClan = generateClannedKeyCheckClan()
            AddChat vbRed, "Failed to check clanned key. Trying again..."
      
            frmMain.sckClanMembers(index).Close
            frmMain.tmrReconnect(index).Enabled = True
        Case &H2
            AddChat vbRed, "Initiate #" & index & ": This key is clanned. Reconnecting on new key."
            addKey bot(index).keyIndex, bot(index).key, KeyType.CLANNED
      
            bot(index).key = keys.getKey(keyIdx)
            bot(index).keyIndex = keyIdx
      
            If (bot(index).key = vbNullString) Then
                AddChat vbRed, "Initiate #" & index & ": No more keys available."
                resetAll
            Else
                frmMain.sckClanMembers(index).Close
                frmMain.tmrReconnect(index).Enabled = True
            End If
        Case &H8
            AddChat vbRed, "Initiate #" & index & ": This initiate is part of a clan. Reconnecting with a new name..."
      
            Dim acc As New clsAccount
            Set acc = generateInitiate()
      
            bot(index).username = acc.getUsername()
            bot(index).password = acc.getPassword()
    
            frmMain.sckClanMembers(index).Close
            frmMain.tmrReconnect(index).Enabled = True
    End Select
End Sub

Public Sub Recv0x72(index As Integer)
    Dim cookie As Long, clanTag As String
    Dim clanName As String, inviterName As String
  
    With Packet(index)
        cookie = .GetDWORD
        clanTag = .GetNonNTString(4)
        clanName = .getNTString
        inviterName = .getNTString
    End With

    With Packet(index)
        .InsertDWORD cookie
        .InsertNonNTString clanTag
        .InsertNTString inviterName
        .InsertByte &H6
        .sendPacket &H72
    End With
End Sub

Public Sub Recv0x0F(index As Integer)
    Dim ID As Long, text As String
    
    ID = Packet(index).GetDWORD
    
    With Packet(index)
        .Skip 20
        .getNTString
        text = .getNTString
    End With
    
    If (ID = &H7) Then
        If (text <> config.Channel) Then
            bot(index).hasRestrictedKey = True
        End If

        Send0x65 index
    End If
End Sub

Public Sub Send0x0A(index As Integer)
    With Packet(index)
        .InsertNTString bot(index).username
        .InsertByte &H0
        .sendPacket &HA
    End With
End Sub

Public Sub Send0x0C(index As Integer)
    With Packet(index)
        .InsertDWORD &H2
        .InsertNTString config.Channel
        .sendPacket &HC
    End With
End Sub

Public Sub Recv0x0A(index As Integer)
    AddChat vbGreen, "Initiate #" & index & ": Logged into Battle.Net!"
    
    Send0x0C index
End Sub

