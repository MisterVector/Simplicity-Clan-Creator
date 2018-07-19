Attribute VB_Name = "modConfig"
Public Sub loadConfig()
    On Error Resume Next
  
    Dim tVal As Variant
  
    tVal = ReadINI("Window", "RememberWindowPosition", "Config.ini")
    config.rememberWindowPosition = (UCase(tVal) = "Y")
  
    If (tVal = vbNullString) Then
        config.rememberWindowPosition = DEFAULT_REMEMBER_WINDOW_POSITION
    End If
  
    If (config.rememberWindowPosition) Then
        frmMain.chkRememberWindowPosition.value = 1
    End If
  
    tVal = ReadINI("Window", "Top", "Config.ini")
  
    If (IsNumeric(tVal) And tVal > 0) Then
        config.windowTop = tVal
    End If
  
    If (config.windowTop = 0) Then
        config.windowTop = frmMain.Top
    End If
  
    tVal = ReadINI("Window", "Left", "Config.ini")
  
    If (IsNumeric(tVal) And tVal > 0) Then
        config.windowLeft = tVal
    End If
  
    If (config.windowLeft = 0) Then
        config.windowLeft = frmMain.Left
    End If
  
    config.server = ReadINI("Main", "Server", "Config.ini")
  
    If (config.server = vbNullString) Then
        config.server = DEFAULT_SERVER
    End If
  
    config.bnlsServer = ReadINI("Main", "BNLSServer", "Config.ini")
  
    If (config.bnlsServer = vbNullString) Then
        config.bnlsServer = DEFAULT_BNLS_SERVER
    End If
  
    tVal = "&H" & ReadINI("Main", "VerByte", "Config.ini")
  
    If (IsNumeric(tVal) And tVal > 0) Then
        config.verByte = tVal
    End If
  
    If (config.verByte = 0) Then
        config.verByte = DEFAULT_VERSION_BYTE
    End If
  
    tVal = ReadINI("Main", "ConnectionTimeOut", "Config.ini")
  
    If (IsNumeric(tVal) And tVal > 0) Then
        config.timeOut = tVal
    End If
    
    If (config.timeOut = 0) Then
        config.timeOut = DEFAULT_TIMEOUT
    End If
  
    tVal = ReadINI("Main", "ReconnectTime", "Config.ini")
  
    If (IsNumeric(tVal) And tVal > 0) Then
        config.reconnectTime = tVal
    End If
  
    If (config.reconnectTime = 0) Then
        config.reconnectTime = DEFAULT_RECONNECT_TIME
    End If
  
    tVal = ReadINI("Main", "ConnectionsPerProxy", "Config.ini")
  
    If (IsNumeric(tVal) And tVal > 0) Then
        config.connectsPerProxy = tVal
    End If
  
    If (config.connectsPerProxy = 0) Then
        config.connectsPerProxy = DEFAULT_CONNECTIONS_PER_PROXY
    End If
  
    config.Channel = ReadINI("Main", "Channel", "Config.ini")
  
    If (config.Channel = vbNullString) Then
        config.Channel = DEFAULT_CHANNEL
    End If
  
    tVal = ReadINI("Main", "SaveClanInfo", "Config.ini")
    config.saveClanInfo = (UCase(tVal) = "Y")
  
    If (tVal = vbNullString) Then
        config.saveClanInfo = DEFAULT_SAVE_CLAN_INFO
    End If
  
    If config.saveClanInfo Then
        frmMain.chkSaveClanInfo.value = 1
    End If
  
    tVal = ReadINI("Main", "UseCustomInitiates", "Config.ini")
    config.useCustomInitiates = (UCase(tVal) = "Y")
    
    If (tVal = vbNullString) Then
        config.useCustomInitiates = DEFAULT_USE_CUSTOM_INITIATES
    End If
  
    If config.useCustomInitiates Then
        frmMain.chkCustomInitiates.value = 1
    End If
    
    tVal = ReadINI("Main", "CheckUpdateOnStartup", "Config.ini")
    config.checkUpdateOnStartup = (UCase(tVal) = "Y")
    
    If (tVal = vbNullString) Then
        config.checkUpdateOnStartup = DEFAULT_CHECK_UPDATE_ON_STARTUP
    End If
    
    If (config.checkUpdateOnStartup) Then
        frmMain.chkCheckUpdateOnStartup.value = 1
    End If
  
    With frmMain
        .txtTimeOut.text = config.timeOut
        .txtPerProxy.text = config.connectsPerProxy
        .txtReconnectTime.text = config.reconnectTime
        .txtChannel.text = config.Channel
        .cmbServer.text = config.server
        .txtBNLSServer.text = config.bnlsServer
        
        .txtChief = ReadINI("Main", "Chieftain", "Config.ini")
        .txtChiefPass = ReadINI("Main", "ChieftainPassword", "Config.ini")
        .txtClanTag = ReadINI("Main", "ClanTag", "Config.ini")
        .txtClanDescription = ReadINI("Main", "ClanDescription", "Config.ini")
        .txtInitiate = ReadINI("Main", "Initiate", "Config.ini")
        .txtInitiatesPassword = ReadINI("Main", "InitiatePass", "Config.ini")
    End With

    If err.Number > 0 Then
        err.Clear
    
        MsgBox "Errors were encountered while loading. The affected values have been set to their defaults.", vbOKOnly Or vbExclamation, PROGRAM_TITLE
    End If
End Sub

Public Sub saveConfig(ByVal saveClanInfo As Boolean)
    WriteINI "Window", "RememberWindowPosition", IIf(config.rememberWindowPosition, "Y", "N"), "Config.ini"
    WriteINI "Window", "Top", config.windowTop, "Config.ini"
    WriteINI "Window", "Left", config.windowLeft, "Config.ini"
    WriteINI "Main", "Server", config.server, "Config.ini"
    WriteINI "Main", "BNLSServer", config.bnlsServer, "Config.ini"
    WriteINI "Main", "ConnectionsPerProxy", config.connectsPerProxy, "Config.ini"
    WriteINI "Main", "ConnectionTimeOut", config.timeOut, "Config.ini"
    WriteINI "Main", "ReconnectTime", config.reconnectTime, "Config.ini"
    WriteINI "Main", "Channel", config.Channel, "Config.ini"
    WriteINI "Main", "SaveClanInfo", IIf(config.saveClanInfo, "Y", "N"), "Config.ini"
    WriteINI "Main", "UseCustomInitiates", IIf(config.useCustomInitiates, "Y", "N"), "Config.ini"
    WriteINI "Main", "CheckUpdateOnStartup", IIf(config.checkUpdateOnStartup, "Y", "N"), "Config.ini"

    If saveClanInfo Then
        WriteINI "Main", "Chieftain", chief.username, "Config.ini"
        WriteINI "Main", "ChieftainPassword", chief.password, "Config.ini"
        WriteINI "Main", "ClanTag", frmMain.txtClanTag.text, "Config.ini"
        WriteINI "Main", "ClanDescription", frmMain.txtClanDescription.text, "Config.ini"
        WriteINI "Main", "Initiate", config.initiate, "Config.ini"
        WriteINI "Main", "InitiatePass", config.initiatePassword, "Config.ini"
    End If
  
    If initiateManager.countInitiates() > 0 Then
        initiateManager.resetIndex
  
        Open App.path & "\CustomInitiates.txt" For Output As #1
            Dim initiates() As New clsAccount
            initiates = initiateManager.getInitiates()
            
            For i = 0 To UBound(initiates)
                Print #1, initiates(i).getUsername() & "|" & initiates(i).getPassword()
            Next i
        Close #1
    Else
        If getFileSize(App.path & "\CustomInitiates.txt") > 0 Then
            Kill App.path & "\CustomInitiates.txt"
        End If
    End If
End Sub

Public Function loadProxies() As Boolean
    Dim tProxies() As String, arrPF() As Variant
    Dim IP As String, Port As Long, Version As String
  
    modProxy.resetProxies
    arrPF = Array("SOCKS4.txt", "HTTP.txt")
  
    For Each proxyfile In arrPF
        If Dir$(App.path & "\" & proxyfile) = vbNullString Then
            Open App.path & "\" & proxyfile For Output As #1
            Close #1
        End If
    
        If getFileSize(App.path & "\" & proxyfile) > 0 Then
            Open App.path & "\" & proxyfile For Input As #1
                tProxies = Split(Input(LOF(1), 1), vbNewLine)
            Close #1

            For i = 0 To UBound(tProxies)
                If tProxies(i) <> vbNullString Then
                    tProxies(i) = Trim(tProxies(i))
                    
                    If InStr(tProxies(i), ":") Then
                        IP = Split(tProxies(i), ":")(0)
                        Port = Split(tProxies(i), ":")(1)
                
                        If IsNumeric(Port) And IsNumeric(Replace(IP, ".", "")) Then
                            If Port > 0 And Port <= 65535 Then
                                Select Case proxyfile
                                    Case "SOCKS4.txt": Version = "s4"
                                    Case "SOCKS5.txt": Version = "s5"
                                    Case "HTTP.txt": Version = "http"
                                End Select
                
                                modProxy.addProxy IP, Port, Version
                            End If
                        End If
                    End If
                End If
            Next i
        End If
    Next
End Function

Public Function loadCustomInitiates() As Boolean
    Dim username As String, password As String
    Dim currentLine As String
  
    Dim fileSize As Integer
    fileSize = getFileSize(App.path & "\CustomInitiates.txt")
  
    If fileSize > 0 Then
        Open App.path & "\CustomInitiates.txt" For Input As #1
            Do Until EOF(1)
                Line Input #1, currentLine
              
                If InStr(currentLine, "|") Then
                    Dim namePass() As String
                    namePass = Split(currentLine, "|")
              
                    initiateManager.addCustomInitiate namePass(0), namePass(1)
                End If
            Loop
        Close #1
    
        loadCustomInitiates = True
    Else
        loadCustomInitiates = False
    End If
End Function
