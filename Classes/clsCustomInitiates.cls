VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomInitiates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private initiates() As New clsAccount
Private idx As Integer

Public Sub Class_Initialize()
    ReDim initiates(0)
End Sub

Public Function addCustomInitiate(initiate As String, password As String) As Boolean
    For i = 0 To UBound(initiates)
        If (LCase(initiates(i).getUsername()) = LCase(initiate)) Then
            addCustomInitiate = False
            Exit Function
        End If
    Next i

    If (initiates(0).getUsername() <> vbNullString) Then
        ReDim Preserve initiates(UBound(initiates) + 1)
    End If

    With initiates(UBound(initiates))
        .setUsername (initiate)
        .setPassword (password)
    End With

    addCustomInitiate = True
End Function

Public Sub updateInitiate(initiate As String, password As String)
    For i = 0 To UBound(initiates)
        If (LCase(initiates(i).getUsername()) = LCase(initiate)) Then
            initiates(i).setPassword (password)
        End If
    Next i
End Sub

Public Sub resetIndex()
    idx = 0
End Sub

Public Function countInitiates() As Integer
    countInitiates = (UBound(initiates) + 1)
End Function

Public Function getNextInitiate() As clsAccount
    Dim acc As New clsAccount
  
    If (idx = (UBound(initiates) + 1)) Then
        Set getNextInitiate = Nothing
        Exit Function
    End If
  
    With initiates(idx)
        acc.setUsername (initiates(idx).getUsername())
        acc.setPassword (initiates(idx).getPassword())
    End With
  
    idx = idx + 1
    Set getNextInitiate = acc
End Function

Public Sub removeInitiate(initiate As String)
    Dim newAcc() As New clsAccount, found As Boolean, tempIdx As Integer: tempIdx = 0
  
    For i = 0 To UBound(initiates)
        If (LCase(initiates(i).getUsername()) = LCase(initiate)) Then
            found = True
            Exit For
        End If
    Next i
  
    If (found) Then
        If (UBound(initiates) = 0) Then
            initiates(0).setUsername (vbNullString)
            initiates(0).setPassword (vbNullString)
    
            Exit Sub
        End If
  
        ReDim newAcc(UBound(initiates) - 1)
    
        For i = 0 To UBound(initiates)
            If (LCase(initiates(i).getUsername()) <> LCase(initiate)) Then
                With newAcc(tempIdx)
                    .setUsername (initiates(i).getUsername())
                    .setPassword (initiates(i).getPassword())
                End With
      
                tempIdx = tempIdx + 1
            End If
        Next i
    
        initiates = newAcc
    End If
End Sub

Public Function getAccountByName(name As String) As clsAccount
    For i = 0 To UBound(initiates)
        If (LCase(initiates(i).getUsername()) = LCase(name)) Then
            Set getAccountByName = initiates(i)
            Exit Function
        End If
    Next i
  
    Set getAccountByName = Nothing
End Function

Public Function getInitiates() As clsAccount()
    getInitiates = initiates
End Function
