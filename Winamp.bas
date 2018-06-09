Attribute VB_Name = "Winamp"
'Windows API Functions
Public Const WINAMP_PATH As String = "C:\Program Files\Winamp\winamp.exe"

Public Const WM_COMMAND = &H111                     'Used in SendMessage call
Public Const WM_USER = &H400                        'Used in SendMessage call
Public Const vbQuote As String = """"              'Used in shelling to WinAMP

Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMessage As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function IsWindow Lib "User32" (ByVal hWnd As Long) As Long
Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "User32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

'--------------------------------------'
'         User Message Constants       '
'--------------------------------------'
Public Const WA_GETVERSION = 0
Public Const WA_CLEARPLAYLIST = 101
Public Const WA_GETSTATUS = 104
Public Const WA_GETTRACKPOSITION = 105
Public Const WA_GETTRACKLENGTH = 105
Public Const WA_SEEKTOPOSITION = 106
Public Const WA_SETVOLUME = 122
Public Const WA_SETBALANCE = 123
Public Const WA_GETEQDATA = 127
Public Const WA_SETEQDATA = 128
'--------------------------------------'
'      Command Message Constants       '
'--------------------------------------'
Public Const WA_PREVTRACK = 40044
Public Const WA_NEXTTRACK = 40048
Public Const WA_PLAY = 40045
Public Const WA_PAUSE = 40046
Public Const WA_STOP = 40047
Public Const IPC_SETVOLUME = 122
Public Const WA_FADEOUTSTOP = 40147
Public Const WA_STOPAFTERTRACK = 40157
Public Const WA_FASTFORWARD = 40148                 '5 secs
Public Const WA_FASTREWIND = 40144                  '5 secs
Public Const WA_PLAYLISTHOME = 40154
Public Const WA_PLAYLISTEND = 40158
Public Const WA_DIALOGOPENFILE = 40029
Public Const WA_DIALOGOPENURL = 40155
Public Const WA_DIALOGFILEINFO = 40188
Public Const WA_TIMEDISPLAYELAPSED = 40037
Public Const WA_TIMEDISPLAYREMAINING = 40038
Public Const WA_TOGGLEPREFERENCES = 40012
Public Const WA_DIALOGVISUALOPTIONS = 40190
Public Const WA_DIALOGVISUALPLUGINOPTIONS = 40191
Public Const WA_STARTVISUALPLUGIN = 40192
Public Const WA_TOGGLEABOUT = 40041
Public Const WA_TOGGLEAUTOSCROLL = 40189
Public Const WA_TOGGLEALWAYSONTOP = 40019
Public Const WA_TOGGLEWINDOWSHADE = 40064
Public Const WA_TOGGLEPLAYLISTWINDOWSHADE = 40266
Public Const WA_TOGGLEDOUBLESIZE = 40165
Public Const WA_TOGGLEEQ = 40036
Public Const WA_TOGGLEPLAYLIST = 40040
Public Const WA_TOGGLEMAINWINDOW = 40258
Public Const WA_TOGGLEMINIBROWSER = 40298
Public Const WA_TOGGLEEASYMOVE = 40186
Public Const WA_VOLUMEUP = 40058                    'increase 1%
Public Const WA_VOLUMEDOWN = 40059                  'decrease 1%
Public Const WA_TOGGLEREPEAT = 40022
Public Const WA_TOGGLESHUFFLE = 40023
Public Const WA_DIALOGJUMPTOTIME = 40193
Public Const WA_DIALOGJUMPTOFILE = 40194
Public Const WA_DIALOGSKINSELECTOR = 40219
Public Const WA_DIALOGCONFIGUREVISUALPLUGIN = 40221
Public Const WA_RELOADSKIN = 40291
Public Const WA_CLOSE = 40001
Public Const WM_WA_IPC = 1024
Public Sub LoadWinamp()
    Dim strPath$, dblRes As Double
    strPath = "C:\Program Files\Winamp\winamp.exe"
    If Dir$(strPath) = "" Then 'no winamp
    Send0x0E "Can't find Winamp installation!"
    Exit Sub
    End If
On Error Resume Next
    dblRes = Shell(strPath, vbMinimizedNoFocus)
End Sub
''' DOES NOT WORK :(
Public Sub LoadPlaylist()
    Dim strPath$, dblRes As Double
    strPath = "PLAYLISTS\" & pltitle & ".m3u"
On Error Resume Next
    dblRes = Shell(strPath, vbMinimizedNoFocus)
End Sub
Function WinAMP_FindWindow() As Boolean

    '-------------------------------------'
    'Retrieves a handle to WinAMP's window'
    '-------------------------------------'

    hWndWinAMP = FindWindow("Winamp v1.x", vbNullString)
    If hWndWinAMP <> 0 Then
       WinAMP_FindWindow = True
    Else
       WinAMP_FindWindow = False
    End If

End Function

Function WinAMP_GetStatus(ByRef Success As Boolean) As String

    '----------------------------------------------------------'
    'Retrieves the status of WinAMP: PLAYING, PAUSED or STOPPED'
    '----------------------------------------------------------'

    Dim Status As Long
    Dim i As Long

    If hWndWinAMP = 0 Then
        Success = False
        Exit Function
    End If

    Status = SendMessage(hWndWinAMP, WM_USER, 0, WA_GETSTATUS)
    
    Select Case Status
       Case 1
          WinAMP_GetStatus = "Winamp is playing."
       Case 3
          WinAMP_GetStatus = "Winamp is paused."
       Case Else
          WinAMP_GetStatus = "Winamp is stopped."
    End Select

End Function

Function WinAMP_GetVersion() As String

    '---------------------------------------'
    'Retrieves the version of WinAMP running'
    '---------------------------------------'

    Dim VersionNum As Long
    Dim ReturnVersion As String

    If hWndWinAMP = 0 Then
       'MessageBox vbOKOnly + vbCritical, "WinAMP window not found yet...WinAMP Not Found"
       Exit Function
    End If

    VersionNum = SendMessage(hWndWinAMP, WM_USER, 0, WA_GETVERSION)
    
    If Len(Hex(VersionNum)) > 3 Then
       ReturnVersion = Left(Hex(VersionNum), 1) & "."
       ReturnVersion = ReturnVersion & Mid(Hex(VersionNum), 2, 1)
       ReturnVersion = ReturnVersion & Right$(Hex(VersionNum), Len(Hex(VersionNum)) - 3)
       WinAMP_GetVersion = ReturnVersion
    Else
       WinAMP_GetVersion = "UNKNOWN"
    End If

End Function

Function WinAMP_SendCommandMessage(CommandMessage As Long) As Integer
    
    '--------------------------------------------------'
    'Used to send any of the Command Messages to WinAMP'
    '--------------------------------------------------'
    hWndWinAMP = FindWindow("Winamp v1.x", vbNullString)
    If hWndWinAMP = 0 Then
       WinAMP_SendCommandMessage = 1
       Exit Function
    End If
    SendMessage hWndWinAMP, WM_COMMAND, CommandMessage, 0

End Function

Public Function WinAMP_Start() As Boolean
    WinAMP_Start = True
End Function

Function WinAMP_GetTrackPosition() As Long

    '---------------------------------------------------'
    'Retrieves the position of the current track in secs'
    '---------------------------------------------------'

    Dim ReturnPos As Long
    
    If hWndWinAMP = 0 Then
       Exit Function
    End If

    'ReturnPos will contain the current track pos in milliseconds
    'or -1 if no track is playing or an error occurs
    ReturnPos = SendMessage(hWndWinAMP, WM_USER, 0, WA_GETTRACKPOSITION)

    If ReturnPos <> -1 Then
       WinAMP_GetTrackPosition = ReturnPos \ 1000   'convert ReturnPos to secs
    Else
       WinAMP_GetTrackPosition = -1
    End If

End Function

Function GetWindowTitle(ByVal WindowTitle As String) As String
    Dim Title As String
    TheHWnd = FindWindow(WindowTitle, vbNullString)
        Title = Space$(GetWindowTextLength(TheHWnd) + 1)
        Call GetWindowText(TheHWnd, Title, Len(Title))
        Title = Left$(Title, Len(Title) - 1)
    GetWindowTitle = Title
End Function




