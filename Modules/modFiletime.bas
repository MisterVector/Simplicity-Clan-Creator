Attribute VB_Name = "modFiletime"
'*******************************
'******* Battle.Net Hash *******
'*****      Control        *****
'*****      By Punk        *****
'*******************************

'Do not modify this file!
'This is part of BNHash functionality and could possibly be updated. If you don't want to lose anywork
'then it's advised that you create your own module.

Option Explicit

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

Private Declare Function FileTimeToLocalFileTime Lib "Kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "Kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

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

