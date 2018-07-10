Attribute VB_Name = "modBNETAPI"
Private Declare Function FileTimeToLocalFileTime Lib "Kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "Kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Public Type FILETIME
  dwLowDateTime       As Long
  dwHighDateTime      As Long
End Type

Public Type SYSTEMTIME
  wYear               As Integer
  wMonth              As Integer
  wDayOfWeek          As Integer
  wDay                As Integer
  wHour               As Integer
  wMinute             As Integer
  wSecond             As Integer
  wMilliseconds       As Integer
End Type
Public tpLocal As SYSTEMTIME
Public tpSystem As SYSTEMTIME

Public Declare Function nls_init Lib "libbnet.dll" (ByVal sUsername As String, ByVal sPassword As String) As Long
Public Declare Function nls_reinit Lib "libbnet.dll" (ByVal lNLSPointer As Long, ByVal sUsername As String, ByVal sPassword As String) As Long
Public Declare Sub nls_free Lib "libbnet.dll" (ByVal lNLSPointer As Long)
Public Declare Function nls_account_logon Lib "libbnet.dll" (ByVal lNLSPointer As Long, ByVal sBufferOut As String) As Long
Public Declare Sub nls_account_logon_proof Lib "libbnet.dll" (ByVal lNLSPointer As Long, ByVal sBufferOut As String, ByVal sServerKey As String, ByVal sSalt As String)
Public Declare Function nls_account_create Lib "libbnet.dll" (ByVal lNLSPointer As Long, ByVal sBufferOut As String) As Long
Public Declare Function nls_account_change Lib "libbnet.dll" (ByVal lNLSPointer As Long, ByVal sBufferOut As String) As Long
Public Declare Function nls_account_change_proof Lib "libbnet.dll" (ByVal lNLSPointer As Long, ByVal sBufferOut As String, ByVal sNewPassword As String, ByVal sServerKey As String, ByVal sSalt As String) As Long
Public Declare Function nls_account_upgrade_proof Lib "libbnet.dll" (ByVal lNLSPointer As Long, ByVal sBufferOut As String) As Long
Public Declare Function checkRevision_ld Lib "libbnet.dll" Alias "checkrevision_ld" (ByVal sFile1 As String, ByVal sFile2 As String, ByVal sFile3 As String, ByVal sValueString As String, ByRef lVersion As Long, ByRef lChecksum As Long, ByVal sReturnDigest As String, ByVal sLockdownFile As String, ByVal sVideoFile As String) As Long
Public Declare Function checkRevision Lib "libbnet.dll" Alias "checkrevision" (ByVal sFile1 As String, ByVal sFile2 As String, ByVal sFile3 As String, ByVal sValueString As String, ByRef lVersion As Long, ByRef lChecksum As Long, ByVal sExeInfo As String, ByVal sMPQName As String) As Long
Public Declare Function decode_hash_cdkey Lib "libbnet.dll" (ByVal sCDKey As String, ByVal lClientToken As Long, ByVal lServerToken As Long, ByRef lPublicValue As Long, ByRef lProductID As Long, ByVal sBufferOut As String) As Long
Public Declare Function decode_hash_cdkey_36 Lib "libbnet.dll" (ByVal sCDKey As String, ByVal lClientToken As Long, ByVal lServerToken As Long, ByRef lPublicValue As Long, ByRef lProductID As Long, ByVal sBufferOut As String) As Long
Public Declare Sub double_hash_password Lib "libbnet.dll" (ByVal sPassword As String, ByVal lClientToken As Long, ByVal lServerToken As Long, ByVal sBufferOut As String)
Public Declare Sub hash_password Lib "libbnet.dll" (ByVal sPassword As String, ByVal sBufferOut As String)

Public Declare Function check_revision Lib "VersionCheck.dll" (ByVal archiveTime As String, ByVal ArchiveName As String, ByVal Seed As String, ByVal INIFile As String, ByVal INIHeader As String, ByRef Version As Long, ByRef Checksum As Long, ByVal result As String) As Long
Public Declare Function crev_max_result Lib "VersionCheck.dll" () As Long

Public Function GetFTTime(FT As FILETIME, Optional Shorten As Boolean = False, Optional localTime As Boolean = True) As String
    Dim LocalFT As FILETIME
    Dim SysTime As SYSTEMTIME
    Dim SetHour As String
    Dim AP      As String

    If localTime Then
        FileTimeToLocalFileTime FT, LocalFT
        FileTimeToSystemTime LocalFT, SysTime
    Else
        FileTimeToSystemTime FT, SysTime
    End If
  
    If SysTime.wHour = 0 Then
        AP = "AM"
        SetHour = "12"
    ElseIf SysTime.wHour < 12 Then
        AP = "AM"
        SetHour = Trim$(str$(SysTime.wHour))
    ElseIf SysTime.wHour = 12 Then
        AP = "PM"
        SetHour = "12"
    Else
        AP = "PM"
        SetHour = Trim$(str$(SysTime.wHour))
    End If
  
    SysTime.wDayOfWeek = SysTime.wDayOfWeek + 1
    
    If Shorten Then
        GetFTTime = Format$(SysTime.wMonth, "00") & "/" & Format$(SysTime.wDay, "00") & "/" & Right$(SysTime.wYear, 2) & " " & SetHour & ":" & Format$(SysTime.wMinute, "00") & ":" & Format$(SysTime.wSecond, "00") & " " & AP
    Else
        'GetFTTime = ConvertShortToLong(WeekdayName(SysTime.wDayOfWeek, True)) & ", " & MonthName(SysTime.wMonth, True) & " " & SysTime.wDay & ", " & SysTime.wYear & " at " & SetHour & ":" & Format$(SysTime.wMinute, "00") & ":" & Format$(SysTime.wSecond, "00") & " " & AP
    End If
End Function

Private Function ConvertShortToLong(Day As String)
    Select Case Day
        Case "Mon": ConvertShortToLong = "Monday"
        Case "Tue": ConvertShortToLong = "Tuesday"
        Case "Wed": ConvertShortToLong = "Wednesday"
        Case "Thu": ConvertShortToLong = "Thursday"
        Case "Fri": ConvertShortToLong = "Friday"
        Case "Sat": ConvertShortToLong = "Saturday"
        Case "Sun": ConvertShortToLong = "Sunday"
    End Select
End Function

