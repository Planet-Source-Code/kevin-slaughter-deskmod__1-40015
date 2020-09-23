Attribute VB_Name = "modMisc"
'API
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, ptry As NOTIFYICONDATA) As Boolean

' Notify icon stuff
Public Type NOTIFYICONDATA
     cbSize As Long
     hwnd As Long
     uId As Long
     uFlags As Long
     uCallBackMessage As Long
     hIcon As Long
     szTip As String * 64
End Type
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

'Mouse
Public Type POINTAPI
    x As Long
    y As Long
End Type

'Force-Refresh
Public Const WM_PAINT = &HF

'Menu stuff
Public Type RECT
    bottom As Long
    left As Long
    right As Long
    top As Long
End Type
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
Public Const MF_STRING = &H0&
Public Const MF_SEPARATOR = &H800&
Public Const MF_CHECKED = &H8&


'// INI Files
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function ReadINI(sFile As String, Section_Title As String, sKeyName As String) As String
Dim Ret As String, NC As Long
    Ret = String(255, 0)
    NC = GetPrivateProfileString(Section_Title, sKeyName, "ERR", Ret, 255, sFile)
    If NC <> 0 Then ReadINI = left$(Ret, NC)
End Function

Public Function WriteINI(FileName As String, Section_Title As String, KeyName As String, ValData As String)
    WritePrivateProfileString Section_Title, KeyName, ValData, FileName
End Function


'// Tiles are not available outside of WinXP, so we use this
'//   to easily check what version is being run.
Public Function IsWinXP() As Boolean
    Dim GetWinVersion As String, WinVer As Long
    WinVer = GetVersion() And &HFFFF&
    IsWinXP = CBool(Format((WinVer Mod 256) + ((WinVer \ 256) / 100), "Fixed") >= 5.01)
End Function

