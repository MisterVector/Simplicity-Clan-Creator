Attribute VB_Name = "modChiefBNET"
Public Sub Chief_Send0x00()
    chiefPacketHandler.sendPacket &H0
End Sub

Public Sub Chief_Recv0x00()
    Chief_Send0x00
End Sub

Public Sub Chief_Send0x25()
    With chiefPacketHandler
        .InsertDWORD .GetDWORD
        .sendPacket &H25
    End With
End Sub

Public Sub Chief_Recv0x25()
    Chief_Send0x25
End Sub

Public Sub Chief_Send0x50()
    With chiefPacketHandler
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

Public Sub Chief_Recv0x50()
    Dim clientToken As Long, serverToken As Long
    Dim MPQFileN As String, valueString As String
    Dim FT As FILETIME, archiveTime As String

    chiefPacketHandler.Skip 4              'Logon type
  
    clientToken = GetTickCount
    serverToken = chiefPacketHandler.GetDWORD
    chiefPacketHandler.Skip 4    'UDPValue
  
    FT.dwLowDateTime = chiefPacketHandler.GetDWORD
    FT.dwHighDateTime = chiefPacketHandler.GetDWORD
    archiveTime = GetFTTime(FT)

    MPQFileN = chiefPacketHandler.getNTString
    valueString = chiefPacketHandler.getNTString
  
    Chief_Send0x51 clientToken, serverToken, MPQFileN, valueString, archiveTime
End Sub

Public Sub Chief_Send0x51(clientToken As Long, serverToken As Long, MPQFileN As String, valueString As String, archiveTime As String)
    Dim EXEInfo      As String: EXEInfo = String$(crev_max_result, Chr$(0))
    Dim Checksum     As Long
    Dim Version      As Long
    Dim CDKeyHash As String * 20
    Dim ProdVal   As Long
    Dim PubVal    As Long
  
    If (decode_hash_cdkey(chief.key, clientToken, serverToken, PubVal, ProdVal, CDKeyHash) = 0) Then
        MsgBox "Unable to decode your CD-Keys.", vbOKOnly Or vbCritical, PROGRAM_TITLE
        Unload All
    
        Exit Sub
    End If

    check_revision archiveTime, MPQFileN, valueString, App.path & "\VersionCheck.ini", "WAR3", Version, Checksum, EXEInfo
  
    With chiefPacketHandler
        .InsertDWORD clientToken
        .InsertDWORD Version
        .InsertDWORD Checksum
        .InsertDWORD &H1
        .InsertDWORD &H0
        
        .InsertDWORD Len(chief.key)
        .InsertDWORD ProdVal
        .InsertDWORD PubVal
        .InsertDWORD &H0
        .InsertNonNTString CDKeyHash
        
        .InsertNTString KillNull(EXEInfo)
        .InsertNTString "Simplicity"
        .sendPacket &H51
    End With
End Sub

Public Sub Chief_Recv0x51()
    Dim statusCode As Long, keyIdx As Long
  
    frmMain.tmrChiefTimeout.Enabled = False
    statusCode = chiefPacketHandler.GetDWORD
  
    If (statusCode = &H0) Then
        Chief_Send0x53
    Else
        frmMain.sckChieftain.Close
  
        Select Case statusCode
            Case &H100, &H101
                resetAll True
        
                If (attemptedVerByteUpdate) Then
                    AddChat vbRed, "Warcraft III hashes are out of date. You need to update them."
                Else
                    AddChat vbRed, "Chieftain: Your game version is out of date."
                    AddChat vbYellow, "Attempting to update version byte..."
          
                    If (config.bnlsServer <> vbNullString) Then
                        frmMain.sckBNLS.Connect config.bnlsServer, 9367
                    Else
                        AddChat vbRed, "BNLS server not configured. Unable to update version byte!"
                    End If
                End If
        
                Exit Sub
            Case &H102
                AddChat vbRed, "Chieftain: Your game version must be downgraded."
                resetAll
                Exit Sub
            Case &H201
                AddChat vbRed, "Chieftain: Your CD-Key is in use by " & chiefPacketHandler.getNTString & "."
                addKey chief.keyIndex, chief.key, KeyType.IN_USE
            Case &H202, &H203
                If (statusCode = &H202) Then
                    AddChat vbRed, "Chieftain: That CD-Key is banned."
                Else
                    AddChat vbRed, "Chieftain: That CD-Key has the wrong product."
                End If
        
                addKey chief.keyIndex, chief.key, KeyType.BAD
        End Select
    
        chief.key = keys.getKey(keyIdx)
        chief.keyIndex = keyIdx
    
        If (chief.key = vbNullString) Then
            AddChat vbRed, "Chieftain: You have run out of keys."
            resetAll
        Else
            frmMain.tmrChiefReconnect.Enabled = True
        End If
    End If
End Sub

Public Sub Chief_Send0x53()
    Dim nls_A As String

    chief.nls_P = nls_init(chief.username, chief.password)

    If (chief.nls_P = 0) Then
        MsgBox "NLS made a bad call.", vbOKOnly Or vbCritical, PROGRAM_TITLE
        unloadAll
        Exit Sub
    End If

    nls_A = Space$(Len(chief.username) + 33)
    
    If (nls_account_logon(chief.nls_P, nls_A) = 0) Then
        MsgBox "Unable to create NLS key.", vbOKOnly Or vbCritical, PROGRAM_TITLE
    
        unloadAll
        Exit Sub
    End If

    chiefPacketHandler.InsertNonNTString Left$(nls_A, Len(nls_A) - Len(chief.username) - 1)
    chiefPacketHandler.InsertNTString chief.username
    chiefPacketHandler.sendPacket &H53
End Sub

Public Sub Chief_Recv0x53()
    Select Case chiefPacketHandler.GetDWORD
        Case &H0: Chief_Send0x54   'Passed
        Case &H1: Chief_Send0x52   'Account Not made
        Case Else
            AddChat vbYellow, "Chieftain: Account error..."
            resetAll
    End Select
End Sub

Public Sub Chief_Send0x52()
    Dim SaltHash As String: SaltHash = Space$(Len(chief.username) + 65)

    If (nls_account_create(chief.nls_P, SaltHash) = 0) Then
        MsgBox "Unable to create NLS salt.", vbOKOnly Or vbCritical, PROGRAM_TITLE
        Exit Sub
    End If

    chiefPacketHandler.InsertNonNTString SaltHash
    chiefPacketHandler.sendPacket &H52
End Sub

Public Sub Chief_Recv0x52()
    Dim result As Long
  
    result = chiefPacketHandler.GetDWORD
  
    If (result = &H0) Then
        Chief_Send0x53
        AddChat vbGreen, "Chieftain: Account successfully created."
    Else
        frmMain.sckChieftain.Close
        AddChat vbRed, "Chieftain: Account creation error. Retrying..."
        frmMain.tmrChiefReconnect.Enabled = True
    End If
End Sub

Public Sub Chief_Send0x54()
    Dim ProofHash As String * 20
    Dim Salt      As String: Salt = chiefPacketHandler.GetNonNTString(32)
    Dim ServerKey As String: ServerKey = chiefPacketHandler.GetNonNTString(32)

    nls_account_logon_proof chief.nls_P, ProofHash, ServerKey, Salt

    chiefPacketHandler.InsertNonNTString ProofHash
    chiefPacketHandler.sendPacket &H54
End Sub

Public Sub Chief_Recv0x54()
    Select Case chiefPacketHandler.GetDWORD
        Case &H0: GoTo Continue
        Case &H1:
        Case &H2: AddChat vbRed, "Chieftain: Invalid password for " & chief.username & "."
        Case &HF:
        Case &HE:
            GoTo Continue
    End Select
    
    resetAll
    Exit Sub
Continue:

    nls_free (chief.nls_P)

    Chief_Send0x0A
End Sub

Public Sub Chief_Send0x65()
    With chiefPacketHandler
        .sendPacket &H65
    End With
End Sub

Public Sub Chief_Recv0x65()
    Dim friendsCount As Integer, curFriend As String
  
    With chiefPacketHandler
        friendsCount = .GetByte

        If (friendsCount > 0) Then
            For i = 0 To friendsCount - 1
                curFriend = .getNTString
                chiefData.addFriend curFriend
        
                .Skip 6
                .getNTString
            Next i
        End If
    End With
  
    Chief_Send0x70
End Sub

Public Sub Chief_Send0x70()
    Dim sendTag As String
  
    For i = 0 To 3
        sendTag = sendTag & Chr$(Asc(Mid$(clannedKeyCheckClan, i + 1, 1)))
    Next i
  
    With chiefPacketHandler
        .InsertDWORD &H0
        .InsertNonNTString sendTag
        .sendPacket &H70
    End With
End Sub

Public Sub Chief_Recv0x70()
    Dim result As Long, count As Integer, keyIdx As Long
    Dim tempInitiateList As New Dictionary
    Dim initiateCount As Byte, retryTagCheck As Boolean
    
    chiefPacketHandler.Skip 4
    result = chiefPacketHandler.GetByte
    initiateCount = chiefPacketHandler.GetByte
  
    Select Case result
        Case &H0
            If (isCheckingClanTag) Then
                If (initiateCount < 9) Then
                    Dim acceptedInitiates() As String, reconnectCount As Integer

                    ReDim acceptedInitiates(initiateCount)

                    For i = 1 To initiateCount
                        acceptedInitiates(i) = chiefPacketHandler.getNTString
                    Next i

                    For i = 0 To UBound(bot)
                        Dim found As Boolean
                        found = False

                        For j = 0 To UBound(acceptedInitiates)
                            If (LCase$(bot(i).username) = LCase$(acceptedInitiates(j))) Then
                                found = True
                                Exit For
                            End If
                        Next j

                        If (Not found) Then
                            AddChat vbYellow, "Initiate #" & i & " lost connection. It will be reconnected."
              
                            disconnectInitiate i
                            reconnectInitiate i
              
                            reconnectCount = reconnectCount + 1
                        End If
                    Next i
        
                    If (reconnectCount = 0) Then
                        AddChat vbRed, "There was an error in creating the clan. Please try again."
            
                        resetAll
                    End If
                Else
                    dicInitiatesAdded.RemoveAll
        
                    Dim init As String
        
                    For i = 1 To initiateCount
                        init = chiefPacketHandler.getNTString
            
                        If (isInitiate(init)) Then
                            dicInitiatesAdded.Add init, init
                        End If
                    Next i
        
                    AddChat vbGreen, "Clan " & config.clanTag & " is available!"
                    AddChat vbGreen, "Press ""Create Clan!"" to create this clan!"
                    frmMain.btnCreateClan.Enabled = True
                End If
            Else
                AddChat vbGreen, "Chieftain: Found a non-clanned " & IIf(chief.hasRestrictedKey, "restricted ", vbNullString) & "key!"
                chief.isReadyForPreparation = True
        
                connectedCount = connectedCount + 1
                frmMain.lblConnected.Caption = "Connected: " & Right$(" " & connectedCount, 2)
        
                Dim readyCount As Integer
                readyCount = countReadyForPreparation()

                If (readyCount = 10) Then
                    prepareInitiatesAndChief
                End If
            End If
        Case &H1
            If (isCheckingClanTag) Then
                AddChat vbRed, "The specified clan tag is taken."
                AddChat vbYellow, "You may change the clan tag and description and check again."
                
                frmMain.txtClanTag.Enabled = True
                frmMain.txtClanDescription.Enabled = True
                frmMain.btnCheckClanTag.Enabled = True
            Else
                clannedKeyCheckClan = generateClannedKeyCheckClan()
                frmMain.sckChieftain.Close
                frmMain.tmrChiefReconnect.Enabled = True
            End If
        Case &H2
            AddChat vbRed, "Chieftain: This key is clanned. Reconnecting on new key."
            frmMain.sckChieftain.Close
            addKey chief.keyIndex, chief.key, KeyType.CLANNED
            
            chief.key = keys.getKey(keyIdx)
            chief.keyIndex = keyIdx
            
            If (chief.key = vbNullString) Then
                AddChat vbRed, "Chieftain: No more keys available."
                resetAll
            Else
                frmMain.tmrChiefReconnect.Enabled = True
            End If
        Case &H8
            AddChat vbRed, "The chieftain is already a part of a clan."
            resetAll
    End Select
End Sub

Public Sub Chief_Recv0x71()
    Dim result As Long

    chiefPacketHandler.GetDWORD 'skip cookie
    result = chiefPacketHandler.GetDWORD
  
    Select Case result
        Case &H0
            AddChat vbGreen, "The clan was successfully created!"
            exportClanToFile
            addKey chief.keyIndex, chief.key, KeyType.CLANNED
      
            chief.key = vbNullString
            chief.keyIndex = 0
      
            AddChat vbGreen, "When you are finished click ""Disconnect All""."
        Case &H1
            AddChat vbRed, "The specified clan tag is taken."
            AddChat vbYellow, "You may change the clan tag and description and check again."
            
            frmMain.txtClanTag.Enabled = True
            frmMain.txtClanDescription.Enabled = True
            frmMain.btnCheckClanTag.Enabled = True
        Case &H5
            AddChat vbRed, "Chieftain not available (not in channel or already in a clan)"
        Case Else
            AddChat vbRed, "The clan was not created! Error code 0x" & IIf(Len(result) > 1, result, "0" & result)
    End Select
  
    If (Not (result = &H0 Or result = &H1)) Then
        resetAll
    End If
End Sub

Public Sub Chief_Recv0x72()
    Dim cookie As Long, clanTag As String
    Dim clanName As String, inviterName As String
  
    With chiefPacketHandler
        cookie = .GetDWORD
        clanTag = .GetNonNTString(4)
        clanName = .getNTString
        inviterName = .getNTString
    End With

    With chiefPacketHandler
        .InsertDWORD cookie
        .InsertNonNTString clanTag
        .InsertNTString inviterName
        .InsertByte &H6
        .sendPacket &H72
    End With
End Sub

Public Sub Chief_Recv0x0F()
    Dim ID As Long, text As String
    
    ID = chiefPacketHandler.GetDWORD
    
    With chiefPacketHandler
        .Skip 20
        .getNTString
        text = .getNTString
    End With
    
    If (ID = &H7) Then
        If (text <> config.Channel) Then
            chief.hasRestrictedKey = True
        End If
        
        Chief_Send0x65
    End If
End Sub

Public Sub Chief_Send0x0A()
    With chiefPacketHandler
        .InsertNTString bot(index).username
        .InsertByte &H0
        .sendPacket &HA
    End With
End Sub

Public Sub Chief_Send0x0C()
    With chiefPacketHandler
        .InsertDWORD &H2
        .InsertNTString config.Channel
        .sendPacket &HC
    End With
End Sub

Public Sub Chief_Recv0x0A()
    AddChat vbGreen, "Chieftain: Logged into Battle.Net!"
  
    Chief_Send0x0C
End Sub

