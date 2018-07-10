Attribute VB_Name = "modDNS"
Private Declare Function DnsQuery Lib "dnsapi" Alias "DnsQuery_A" (ByVal strname As String, ByVal wType As Integer, ByVal fOptions As Long, ByVal pServers As Long, ppQueryResultsSet As Long, ByVal pReserved As Long) As Long
Private Declare Function DnsRecordListFree Lib "dnsapi" (ByVal pDnsRecord As Long, ByVal FreeType As Long) As Long
Private Declare Function lstrlen Lib "Kernel32" (ByVal straddress As Long) As Long
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal pIP As Long) As Long
Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal sAddr As String) As Long

Private Const DnsFreeRecordList         As Long = 1
Private Const DNS_TYPE_A                As Long = &H1
Private Const DNS_QUERY_BYPASS_CACHE    As Long = &H8

Private Type VBDnsRecord
    pNext           As Long
    pName           As Long
    wType           As Integer
    wDataLength     As Integer
    flags           As Long
    dwTel           As Long
    dwReserved      As Long
    prt             As Long
    others(35)      As Byte
End Type

Public Function Resolve(sAddr As String) As String
    If IsNumeric(Replace(sAddr, ".", "")) Then
        Resolve = sAddr
        Exit Function
    End If
    
    Dim pRecord     As Long
    Dim pNext       As Long
    Dim uRecord     As VBDnsRecord
    Dim lPtr        As Long
    Dim vSplit      As Variant
    Dim laServers() As Long
    Dim pServers    As Long
    Dim sName       As String

    If DnsQuery(sAddr, DNS_TYPE_A, DNS_QUERY_BYPASS_CACHE, pServers, pRecord, 0) = 0 Then
        pNext = pRecord
        
        Do While pNext <> 0
            Call CopyMemory(uRecord, pNext, Len(uRecord))
            
            If uRecord.wType = DNS_TYPE_A Then
                lPtr = inet_ntoa(uRecord.prt)
                sName = String(lstrlen(lPtr), 0)
                Call CopyMemory(ByVal sName, lPtr, Len(sName))
                
                If LenB(Resolve) <> 0 Then
                    Resolve = Resolve & " "
                End If
                
                Resolve = Resolve & sName
            End If
            
            pNext = uRecord.pNext
        Loop
        
        Call DnsRecordListFree(pRecord, DnsFreeRecordList)
    End If
End Function




