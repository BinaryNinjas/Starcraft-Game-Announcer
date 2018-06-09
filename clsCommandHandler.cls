VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommandHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub ParseCommand(ByVal User As String, ByVal Text As String)
    Dim strArgs() As String, highAccess As Integer
     Dim iwinamp As Integer
    strArgs = Split(Text, " ")
    highAccess = Database.HighestAccess(User)
    If highAccess >= 1 Then
        Select Case PrepareCheck(strArgs(0))
            Case "kick"
                strArgs = Split(Text, " ", 3)
                If UBound(strArgs) = 2 Then
                    HandleKick strArgs(1), strArgs(2)
                Else
                    HandleKick strArgs(1)
                End If
            Case "ban"
                strArgs = Split(Text, " ", 3)
                If UBound(strArgs) = 2 Then
                    HandleBan strArgs(1), strArgs(2)
                Else
                    HandleBan strArgs(1)
                End If
            Case "unban"
                strArgs = Split(Text, " ", 2)
                HandleUnban strArgs(1)
            Case "find"
                strArgs = Split(Text, " ", 2)
                HandleFind strArgs(1)
            Case "findflag"
                strArgs = Split(Text, " ", 2)
                HandleFindFlag strArgs(1)
            Case "cq", "clearq", "clearqueue", "cqueue"
                HandleClearQueue False
            Case "ver", "version", "about"
                HandleVersion
            Case "trigger", "trig", "t"
                HandleTrigger User
            Case "timefilter", "time"
                 If UBound(strArgs) = 0 Then 'blank time
                 Send0x0E "You must input how many seconds you wanted filtered!"
                 Exit Sub
                 End If
                If strArgs(1) = "" Then 'user forgot something
                Send0x0E "You must input how many seconds you wanted filtered!"
                Exit Sub
                End If
            TimeFilter = Trim(strArgs(1))
            Send0x0E "Detecting games within " & TimeFilter & " secs."
            
            Case "clearall", "clear"
            Call ClearAll
            Case "deletefilter", "delete", "del"
            Call DeleteFilter(Trim(LCase(Mid$(Text, Len(strArgs(0)) + 2))))
            
            Case "switch", "watch"
                Select Case UCase(Trim(strArgs(1)))
                Case "UMS"
                WatchType = "UMS"
                Case "TVB"
                WatchType = "TVB"
                Case "MELEE"
                WatchType = "MELEE"
                Case Else
                Send0x0E "Improper Watch Type, use : UMS, TVB, or MELEE"
                Exit Sub
                End Select
                             Send0x0E "Switched to detecting " & WatchType & " games."
   
            Case "notify" 'whisperlist
            Call Whizperlist(strArgs(1))
            Case "track"
            HandleTrack Mid$(Text, 7)
            
            Case "ad", "advertise"
            CreateGame_0x1C Trim(Mid$(Text, Len(strArgs(0)) + 2)), ""
            Send0x0E "Attempting to create " & Trim(Mid$(Text, Len(strArgs(0)) + 2))
            Case "filtercount"
            Send0x0E "You have " & GameFilters2.count & " Filters. {" & Gamefilter.count & " Hashs)"
            
            
            Case "stop", "sleep", "stoptrack"
                    
                    If TimerOn = True Then
                    Call KillTimer(frmMain.hWnd, 1000)
                    Queue.Enqueue "Stopped Tracking!"
                    TimerOn = False
                Else
                    Queue.Enqueue "Tracking is already off"
                    End If
            
            
            Case "start", "starttrack"
                    If TimerOn = False Then
                    Call SetTimer(frmMain.hWnd, 1000, 10000, AddressOf TimerProc)
                    Queue.Enqueue "Tracking Started"
                Else
                    Queue.Enqueue "Tracking is already on"
                    End If
                    
           
  
  
            Case PrepareCheck(IniName)
            Case Else
        End Select
    End If
    If highAccess >= 2 Then
        Select Case PrepareCheck(strArgs(0))
            Case "say"
                strArgs = Split(Text, " ", 2)
                HandleSay strArgs(1)
            Case "recon", "reconnect"
                HandleRecon
            Case "rem", "remove"
             strArgs = Split(Text, " ", 2)
             
            If Database.RemoveDatabase(strArgs(1)) = True Then
            Queue.Enqueue "Removed " & strArgs(1) & " from database"
            Else
            Queue.Enqueue "Can not remove " & strArgs(1)
            End If
            
            Case "add"
                strArgs = Split(Text, " ", 3)
                HandleAdd User, strArgs(1), strArgs(2)
            Case "join"
                strArgs = Split(Text, " ", 2)
                HandleJoin strArgs(1)
            Case "op", "giveop", "giveops"
                strArgs = Split(Text, " ", 2)
                HandleOp strArgs(1)
            Case "resign"
                HandleResign
            
            Case Else
        End Select
    End If
    If highAccess >= 3 Then
        Select Case PrepareCheck(strArgs(0))
            'Case "quit"
            '    HandleQuit
            Case "sethome"
                strArgs = Split(Text, " ", 2)
                HandleSetHome User, strArgs(1)
            Case "settrigger"
                strArgs = Split(Text, " ", 2)
                HandleSetTrigger User, strArgs(1)
                     '''''''WINAMP COMMANDS''''''''''''''''''
       
    Case "winamp"
         
            Select Case strArgs(1)
            
            Case "on"
                LoadWinamp
                WinAMP_SendCommandMessage (WA_PLAY)
                Send0x0E "..::Winamp Loaded::.. [Game Announcer]"
                
            Case "off"
                iwinamp = WinAMP_SendCommandMessage(WA_CLOSE)
                Send0x0E "..::Winamp Closed::.. [Game Announcer]"
            Case "stop"
              
                iwinamp = WinAMP_SendCommandMessage(WA_STOP)
                If i = 0 Then
                Send0x0E "Winamp Stopped"
                Else
                Send0x0E "Winamp is not active"
                End If
  
            End Select
         
    Case "mp3", "music"
            WindowTitle = GetWindowTitle("Winamp v1.x")
            If WindowTitle = vbNullString Then
                Send0x0E "Winamp is not active"
            Else
            WindowTitle = Left(WindowTitle, Len(WindowTitle) - 9)
                Send0x0E "Track [" & WindowTitle & "] ..::Game Announcer::.."
                
            End If
   Case "play", "p"
            iwinamp = WinAMP_SendCommandMessage(WA_PLAY)
            If i = 0 Then
                 WindowTitle = GetWindowTitle("Winamp v1.x")
            If WindowTitle = vbNullString Then
                Send0x0E "Winamp is not active"
            Else
            WindowTitle = Left(WindowTitle, Len(WindowTitle) - 9)
               ' Send0x0E "Track [" & WindowTitle & "] ..::Game Announcer::.."
            End If
                
            Else
                Send0x0E ("Winamp is not active")
            End If
            
                Case "next", "n", "skip"
            iwinamp = WinAMP_SendCommandMessage(WA_NEXTTRACK)
            If i = 0 Then
                WindowTitle = GetWindowTitle("Winamp v1.x")
            If WindowTitle = vbNullString Then
                Send0x0E "Winamp is not active"
            Else
            WindowTitle = Left(WindowTitle, Len(WindowTitle) - 9)
               ' Send0x0E "Track [" & WindowTitle & "] ..::Game Announcer::.."
                
            End If
                 
            Else
                'Send LastCW & "Not on, use /winamp .", frmMain.wsBnet
                Send0x0E "Winamp is not active."
            End If
            
            Case "prev", "back", "last"
            iwinamp = WinAMP_SendCommandMessage(WA_PREVTRACK)
            If i = 0 Then
                WindowTitle = GetWindowTitle("Winamp v1.x")
            If WindowTitle = vbNullString Then
                Send0x0E "Winamp is not active"
            Else
            WindowTitle = Left(WindowTitle, Len(WindowTitle) - 9)
                'Send0x0E "Track [" & WindowTitle & "] ..::Game Announcer::.."
                End If
            Else
                Send0x0E "Winamp is not active"
            End If
            
               Case "pause"
            iwinamp = WinAMP_SendCommandMessage(WA_PAUSE)
            If i = 0 Then
               ' Send0x0E "Winamp Paused"
                
            Else
                Send0x0E "Winamp is not active"
            End If
            
            Case Else
        End Select
    End If
End Sub
Private Sub HandleSay(Text As String)
    Queue.Enqueue Text
End Sub
Private Sub HandleTrigger(strUser As String, Optional intCmd As Boolean = False)
    If intCmd Then
        
        Chat.AddChat vbCyan, "Bot's Trigger: " & IniTrigger
    Else
        Queue.Enqueue "/w " & strUser & " Trigger: " & IniTrigger
    End If
End Sub
Private Sub HandleRecon(Optional intCmd As Boolean = False)
    If intCmd Then
        
        Chat.AddChat vbCyan, "Reconnecting..."
    End If
    frmMain.sckBnet.Close
    frmMain.sckBnet.Connect
End Sub
Private Sub HandleAdd(strUser As String, strAdd As String, strFlags As String, Optional intCmd As Boolean = False)
    If (Database.AddDatabase(strAdd, strFlags) = True) Then
        If intCmd Then
            
            Chat.AddChat vbCyan, "Added " & strAdd & " with flags of " & strFlags & "."
        Else
            Queue.Enqueue "/w " & strUser & " Added " & strAdd & " with flags of " & strFlags & "."
        End If
    Else
        If intCmd Then
            
            Chat.AddChat vbCyan, "Add Failed."
        End If
    End If
End Sub

Private Sub HandleJoin(strChan As String)
    Queue.Enqueue "/join " & strChan
End Sub
Private Sub HandleBan(struserTag As String, Optional banMessage As String = "")
    If Database.CheckSafelist(struserTag) = False Then
        If banMessage = "" Then
            Queue.Enqueue "/ban " & struserTag
        Else
            Queue.Enqueue "/ban " & struserTag & " " & banMessage
        End If
        Else
            Queue.Enqueue struserTag & " is safelisted."
    End If
    
End Sub
Private Sub HandleUnban(strUnban As String)
    Queue.Enqueue "/unban " & strUnban
End Sub
Private Sub HandleKick(struserTag As String, Optional banMessage As String = "")
    If Database.CheckSafelist(struserTag) = False Then
        If banMessage = "" Then
            Queue.Enqueue "/kick " & struserTag
        Else
            Queue.Enqueue "/kick " & struserTag & " " & banMessage
        End If
    Else
        Queue.Enqueue struserTag & " is safelisted."
    End If

End Sub
Private Sub HandleClearQueue(Optional intCmd As Boolean = False)
    Queue.Clear
    If intCmd Then
        
        Chat.AddChat vbCyan, "Queue cleared."
    Else
        Queue.Enqueue "Queue cleared."
    End If
End Sub
Private Sub HandleQuit()
    
    Chat.AddChat vbRed, "Bot Quitting..."
    frmMain.HandleUnload
End Sub
Private Sub HandleSetHome(strUser As String, home As String, Optional intCmd As Boolean = False)
    IniHome = home
    If intCmd Then
        
        Chat.AddChat vbCyan, "Home Channel set to: " & home & "."
    Else
        Queue.Enqueue "/w " & strUser & " Home Channel set to: " & home & "."
    End If
End Sub
Private Sub HandleSetTrigger(strUser As String, trigger As String, Optional intCmd As Boolean = False)
    IniTrigger = trigger
    If intCmd Then
        
        Chat.AddChat vbCyan, "Set trigger to: " & trigger & "."
    Else
        Queue.Enqueue "/w " & strUser & " Set trigger to: " & trigger & "."
    End If
End Sub
Private Sub HandleFind(strFind As String, Optional intCmd As Boolean = False)
    Dim s() As String, i As Integer, tick As Integer, strRet As String
    s = Database.Find(strFind)
    If s(0) = vbNullString Then
        If intCmd Then
            
            Chat.AddChat vbCyan, "No users found."
        Else
            Queue.Enqueue "No users found."
        End If
        Exit Sub
    End If
    For i = 0 To UBound(s)
        If tick >= 10 Then
            tick = 0
            If intCmd Then
                
                Chat.AddChat vbCyan, "Found: " & Left$(strRet, Len(strRet) - 2) & "."
            Else
                Queue.Enqueue "Found: " & Left$(strRet, Len(strRet) - 2) & "."
            End If
            strRet = ""
        End If
        tick = tick + 1
        strRet = strRet & s(i) & ", "
    Next i
    If intCmd Then
        
        Chat.AddChat vbCyan, "Found: " & Left$(strRet, Len(strRet) - 4) & "."
    Else
        Queue.Enqueue "Found: " & Left$(strRet, Len(strRet) - 4) & "."
    End If
End Sub
Private Sub HandleOp(strDes As String, Optional intCmd As Boolean = False)
    Queue.Enqueue "/designate " & strDes
    UserDesignated = strDes
    If intCmd Then
        
        Chat.AddChat vbCyan, "User designated = " & strDes & "."
    Else
        Queue.Enqueue "I have designated: " & strDes
    End If
    Queue.Enqueue "/resign"
End Sub
Private Sub HandleResign()
    Queue.Enqueue "/resign"
End Sub
Private Sub HandleFindFlag(strFindFlag As String, Optional intCmd As Boolean = False)
    Dim s() As String, i As Integer, tick As Integer, strRet As String
    s = Database.FindFlag(strFindFlag)
    If s(0) = vbNullString Then
        If intCmd Then
            
            Chat.AddChat vbCyan, "No users found."
        Else
            Queue.Enqueue "No users found."
        End If
        Exit Sub
    End If
    For i = 0 To UBound(s)
        If tick >= 10 Then
            tick = 0
            If intCmd Then
                
                Chat.AddChat vbCyan, "Found: " & Left$(strRet, Len(strRet) - 2) & "."
            Else
                Queue.Enqueue "Found: " & Left$(strRet, Len(strRet) - 2) & "."
            End If
            strRet = ""
        End If
        tick = tick + 1
        strRet = strRet & s(i) & ", "
    Next i
    If intCmd Then
        
        Chat.AddChat vbCyan, "Found: " & Left$(strRet, Len(strRet) - 4) & "."
    Else
        Queue.Enqueue "Found: " & Left$(strRet, Len(strRet) - 4) & "."
    End If
End Sub
Private Sub HandleVersion(Optional ByVal intCmd As Boolean = False)
    If intCmd Then
        
        Chat.AddChat vbCyan, "Game Announcer v" & App.Major & "." & App.Minor & " (Build: " & App.Revision & ") by Myst and cHip."
    Else
        Queue.Enqueue "Game Announcer v" & App.Major & "." & App.Minor & " (Build: " & App.Revision & ") by Myst and cHip."
    End If
End Sub

