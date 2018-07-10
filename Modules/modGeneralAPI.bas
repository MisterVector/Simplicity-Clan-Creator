Attribute VB_Name = "modGeneralAPI"
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long
Public Declare Function GetCurrentProcess Lib "Kernel32" () As Long
Public Declare Function EmptyWorkingSet Lib "psapi.dll" (ByVal hProcess As Long) As Long
Public Declare Function SetProcessWorkingSetSize Lib "Kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long _
                                                                , ByVal dwMaximumWorkingSetSize As Long) As Long

Public Function ews_memory() As Long:   ews_memory = EmptyWorkingSet(GetCurrentProcess):                  End Function
Public Function spw_memory() As Long:   spw_memory = SetProcessWorkingSetSize(GetCurrentProcess, -1, -1): End Function
Public Sub FreeMemory()
    ews_memory
    spw_memory
End Sub



