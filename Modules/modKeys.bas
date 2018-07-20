Attribute VB_Name = "modKeys"
Public Function loadCDKeys() As Boolean
    Dim arrKeys() As String, tmp As Dictionary

    keys.resetKeys

    Set tmp = New Dictionary

    If (getFileSize(App.path & "\CD-Keys.txt") > 0) Then
        Open App.path & "\CD-Keys.txt" For Input As #1
            arrKeys = Split(Input(LOF(1), 1), vbNewLine)
        Close #1
  
        For i = 0 To UBound(arrKeys)
            arrKeys(i) = UCase(Trim(arrKeys(i)))
    
            If (arrKeys(i) <> vbNullString) Then
                If (Not tmp.Exists(arrKeys(i))) Then
                    tmp.Add arrKeys(i), arrKeys(i)
                End If
            End If
        Next i
  
        For Each key In tmp.keys
            keys.addKey key
        Next
    
        Set tmp = Nothing
        loadCDKeys = True
    End If
End Function

Public Sub addKey(ByVal keyIndex As String, ByVal key As String, kType As KeyType)
    Dim keyFile As String
  
    keys.removeKey keyIndex
  
    Select Case kType
        Case BAD: keyFile = "BadKeys.txt"
        Case IN_USE: keyFile = "InUseKeys.txt"
        Case CLANNED: keyFile = "ClannedKeys.txt"
    End Select
  
    Open App.path & "\" & keyFile For Append As #1
        Print #1, key
    Close #1
End Sub

Public Sub sendBackGoodKeys()
    Dim key As String

    Open App.path & "\CD-Keys.txt" For Output As #1
        For i = 0 To keys.getCount()
            key = keys.getKeyByIndex(i)
    
            If (Not key = vbNullString) Then
                Print #1, key
            End If
        Next i
    Close #1
End Sub

