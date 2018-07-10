Attribute VB_Name = "modBNLS"
Public Sub SEND_BNLS_0x10()
    With bnlsPacket
        .InsertDWORD &H7
        .sendBNLSPacket &H10
    End With
End Sub

Public Sub RECV_BNLS_0x10()
    Dim product As Long

    product = bnlsPacket.GetDWORD
  
    If (product <> 0) Then
        Dim verByte As Long
    
        verByte = bnlsPacket.GetDWORD
  
        config.verByte = verByte
        WriteINI "Main", "VerByte", Hex(verByte), "Config.ini"
    
        attemptedVerByteUpdate = True
    
        AddChat vbGreen, "Updated version byte to 0x" & Hex(verByte) & "!"
        AddChat vbGreen, "You may now click ""Check Keys"" to begin checking again."
    Else
        AddChat vbRed, "Could not update version byte!"
    End If
  
    resetAll True
End Sub
