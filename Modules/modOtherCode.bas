Attribute VB_Name = "modOtherCode"
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

Public Function getBotIndexByName(ByVal user As String) As Integer
    For i = 0 To 8
        If (LCase(bot(i).username) = LCase(user)) Then
            getBotIndexByName = i
            Exit Function
        End If
    Next i
End Function

Public Sub AddChat(ParamArray saElements() As Variant)
    With frmMain.rtbChat
        Dim sp() As String
        sp = Split(.text, vbNewLine)
    
        If (UBound(sp) >= 49) Then
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
  
    If (IsNumeric(Replace(gateway, ".", ""))) Then
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
    If (getFileSize > 0) Then Exit Function
    getFileSize = 0
End Function

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

    If (Not IsProxyPacket) Then
        If (Len(data) >= 12 And LCase(Left(data, 4))) = "http" Then
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

    If (Not IsChiefProxyPacket) Then
        If (Len(data) >= 12 And LCase(Left(data, 4)) = "http") Then
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

Public Function getRealmName(ByVal serverConfig As String) As String
    Dim tempServer As String, IPs() As String, doExit As Boolean
  
    For Each key In dicServerList.keys
        If (LCase(key) = LCase(serverConfig)) Then
            tempServer = key
            Exit For
        End If
    
        IPs = Split(dicServerList.Item(key), "|")
    
        For Each IP In IPs
            If (IP = serverConfig) Then
                tempServer = key
                doExit = True
                Exit For
            End If
        Next
    
        If (doExit) Then
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

Public Function isNewVersion(checkVersion As String) As Boolean
    Dim currentVersionParts() As String, versionParts() As String
    Dim currentVersionPoints As Long, versionPoints As Long
    Dim updated As Boolean
  
    currentVersionParts = Split(PROGRAM_VERSION, ".")
    versionParts = Split(checkVersion, ".")
    
    currentVersionPoints = ((currentVersionParts(0) * 1000000) + (currentVersionParts(1) * 1000) _
                         + currentVersionParts(2))
    
    versionPoints = ((versionParts(0) * 1000000) + (versionParts(1) * 1000) + versionParts(2))

    isNewVersion = (versionPoints > currentVersionPoints)
End Function

