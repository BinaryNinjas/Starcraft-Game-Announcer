Attribute VB_Name = "modAPI"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal numbytes As Long)
Public Declare Function Shell_NotifyIcon Lib "SHELL32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow Lib "User32" (ByVal hWnd As Long) As Long

Public Const BNFLAGS_BLIZZ = &H1
Public Const BNFLAGS_OP = &H2
Public Const BNFLAGS_SPKR = &H4
Public Const BNFLAGS_SYSOP = &H8
Public Const BNFLAGS_PLUG = &H10
Public Const BNFLAGS_MOD = &H16
Public Const BNFLAGS_MOD2 = &H18
Public Const BNFLAGS_SQUELCH = &H20
Public Const BNFLAGS_GLASSES = &H40
Public Const BNFLAGS_GFOFFICIAL = &H100000
Public Const BNFLAGS_GFPLAYER = &H200000

 'creates a region to make form non rectangular
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

'assigns a region to the form
Public Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

 
 

'Used to change the window procedure which kick-starts the subclassing
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Used to call the default window procedure for the parent
Public Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Used to set and retrieve various information
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_NOTIFY = &H4E
Public Const EM_SETEVENTMASK = &H445
Public Const EM_GETEVENTMASK = &H43B
Public Const EM_GETTEXTRANGE = &H44B
Public Const EM_AUTOURLDETECT = &H45B
Public Const EN_LINK = &H70B

Public Const CFE_LINK = &H20
Public Const ENM_LINK = &H4000000
Public Const GWL_WNDPROC = (-4)
Public Const SW_SHOW = 5

Public Const WINSOCK_ERROR = "Windows Sockets not responding correctly."
Public Const INADDR_NONE As Long = &HFFFFFFFF
Public Const WSA_SUCCESS = 0
Public Const WS_VERSION_REQD As Long = &H101

Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Public Declare Function gethostbyname Lib "wsock32" (ByVal hostname As String) As Long
Public Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal addr As Long) As Long
Public Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long

Public Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To 256) As Byte
   szSystemStatus(0 To 128) As Byte
   iMaxSockets As Long
   iMaxUDPDG As Long
   lpVendorInfo As Long
End Type
