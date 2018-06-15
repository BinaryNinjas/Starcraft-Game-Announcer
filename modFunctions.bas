Attribute VB_Name = "modFunctions"
Public TimeFilter As Long
Public GotShowID As Boolean

Public Type UDTLast5Games
GameNames As String
GameTimes As Long
End Type

Public Last5Gamess(4) As UDTLast5Games
Public harvestbool As Boolean
Public HashCnter As Long
Public isMapfilter As Boolean
Public isGamefilter As Boolean
Public isHostfilter As Boolean
Public GameFilters2 As New Collection
Public str As Variant
Public AccessList As New clsHashTable
Public Whisperlist As New clsHashTable
Public GameDatas As New clsHashTable
Public Gamefilter As New clsHashTable

Public Function KillNull(ByVal strText As String) As String
If InStr(strText, Chr(0)) = 0 Then KillNull = strText: Exit Function
    KillNull = Left(strText, InStr(strText, Chr(0)) - 1)
End Function
Public Function getMastFlags() As String
    Dim s As String
    s = Chr(83) & Chr(77)
    getMastFlags = s
End Function
Public Function PrepareCheck(ByVal toCheck As String) As String
    toCheck = Replace(toCheck, "[", "ÿ")
    toCheck = Replace(toCheck, "]", "Ö")
    toCheck = Replace(toCheck, "~", "Ü")
    toCheck = Replace(toCheck, "#", "¢")
    toCheck = Replace(toCheck, "-", "£")
    toCheck = Replace(toCheck, "&", "¥")
    toCheck = Replace(toCheck, "@", "¤")
    toCheck = Replace(toCheck, "{", "ƒ")
    toCheck = Replace(toCheck, "}", "á")
    toCheck = Replace(toCheck, "^", "í")
    toCheck = Replace(toCheck, "`", "ó")
    toCheck = Replace(toCheck, "_", "ú")
    toCheck = Replace(toCheck, "+", "ñ")
    toCheck = Replace(toCheck, "$", "÷")
    PrepareCheck = LCase(toCheck)
End Function
Public Function ConvertTime(ByVal lngMS As Long) As String
    Dim lngSeconds As Long, lngDays As Long, lngHours As Long, lngMins As Long
    Dim strSeconds As String, strDays As String
    lngSeconds = lngMS / 1000
    lngDays = Int(lngSeconds / 86400)
    lngSeconds = lngSeconds Mod 86400
    lngHours = Int(lngSeconds / 3600)
    lngSeconds = lngSeconds Mod 3600
    lngMins = Int(lngSeconds / 60)
    lngSeconds = lngSeconds Mod 60
    If lngSeconds <> 1 Then strSeconds = "s"
    If lngDays <> 1 Then strDays = "s"
    ConvertTime = lngDays & " day" & strDays & ", " & lngHours & " hours, " & lngMins & " minutes and " & lngSeconds & " second" & strSeconds
End Function

Public Function SplitStr(ByVal OriginalString As String, ByRef ReturnArray() As String, ByVal Delimiter As String) As Long
    Dim sItem, lArrCnt
    Dim lLen, lPos
    
    lArrCnt = 0
    lLen = Len(OriginalString)

    Do
        lPos = InStr(1, OriginalString, Delimiter, vbTextCompare)
        If lPos <> 0 Then
            sItem = Left$(OriginalString, lPos - 1)
            OriginalString = Mid$(OriginalString, lPos + 1)
            If sItem <> "" Then
                lArrCnt = lArrCnt + 1
                If lArrCnt = 1 Then
                    ReDim ReturnArray(1 To lArrCnt) As String
                Else
                    ReDim Preserve ReturnArray(1 To lArrCnt) As String
                End If
                ReturnArray(lArrCnt) = sItem
            End If
        End If
    Loop While lPos <> 0

    If OriginalString <> "" Then
        lArrCnt = lArrCnt + 1
        ReDim Preserve ReturnArray(1 To lArrCnt) As String
        ReturnArray(lArrCnt) = OriginalString
    End If
    
    SplitStr = lArrCnt
End Function

Public Function GetIconCode(ByVal Product As String, ByVal Flags As Long) As Integer

    If (BNFLAGS_BLIZZ And Flags) = BNFLAGS_BLIZZ Then
        GetIconCode = 15
        Exit Function
    ElseIf (BNFLAGS_OP And Flags) = BNFLAGS_OP Then
        GetIconCode = 1
        Exit Function
    ElseIf (BNFLAGS_SYSOP And Flags) = BNFLAGS_SYSOP Then
        GetIconCode = 16
        Exit Function
    ElseIf (BNFLAGS_SQUELCH And Flags) = BNFLAGS_SQUELCH Then
        GetIconCode = 14
        Exit Function
    ElseIf (BNFLAGS_GLASSES And Flags) = BNFLAGS_GLASSES Then
        GetIconCode = 18
        Exit Function
    ElseIf (BNFLAGS_GFPLAYER And Flags) = BNFLAGS_GFPLAYER Then
        GetIconCode = 17
        Exit Function
    ElseIf (BNFLAGS_SPKR And Flags) = BNFLAGS_SPKR Then
        GetIconCode = 17
        Exit Function
    End If
    
   Select Case Product
        Case "[SEXP]": GetIconCode = 9
        Case "[STAR]": GetIconCode = 6
        Case "[JSTR]": GetIconCode = 11
        Case "[SSHR]": GetIconCode = 7
        Case "[W2BN]": GetIconCode = 10
        Case "[D2DV]": GetIconCode = 5
        Case "[D2XP]": GetIconCode = 13
        Case "[DRTL]": GetIconCode = 3
        Case "[DSHR]": GetIconCode = 4
        Case "[CHAT]": GetIconCode = 2
        Case "[WAR3]": GetIconCode = WAR3
        Case Else: GetIconCode = 12
    End Select
    
End Function
Public Function GetFlagCode(ByVal Flags As Long) As Integer
    If (BNFLAGS_BLIZZ And Flags) = BNFLAGS_BLIZZ Then
        GetFlagCode = 15
        Exit Function
    ElseIf (BNFLAGS_OP And Flags) = BNFLAGS_OP Then
        GetFlagCode = 1
        Exit Function
    ElseIf (BNFLAGS_SYSOP And Flags) = BNFLAGS_SYSOP Then
        GetFlagCode = 16
        Exit Function
    ElseIf (BNFLAGS_SQUELCH And Flags) = BNFLAGS_SQUELCH Then
        GetFlagCode = 14
        Exit Function
    ElseIf (BNFLAGS_GLASSES And Flags) = BNFLAGS_GLASSES Then
        GetFlagCode = 18
        Exit Function
    ElseIf (BNFLAGS_GFPLAYER And Flags) = BNFLAGS_GFPLAYER Then
        GetFlagCode = 17
        Exit Function
    ElseIf (BNFLAGS_SPKR And Flags) = BNFLAGS_SPKR Then
        GetFlagCode = 17
        Exit Function
    Else
        GetFlagCode = 2
        Exit Function
    End If
End Function
 
Public Function Whizperlist(User As String)

Send0x0E "/f add " & Trim(User)
Send0x0E Trim(User) & " has been added to the whisper list."

End Function
Public Function NotifyAll(Broadcast As String)

Send0x0E "/f msg " & Broadcast
End Function

Public Function AddtoAccess(WhoIniatedCMD As String, User As String, AccessLevel As Integer)

    If AccessList.Exists(LCase(User)) = True Then  'Already In
     If AccessList.Item(LCase(User)) = AccessLevel Then  'then same access
     Chat.AddChat2 vbRed, User & " already has " & AccessLevel
     Exit Function
     End If
     If AccessList.Item(LCase(User)) >= AccessList.Item(LCase(WhoIniatedCMD)) Then
     Chat.AddChat2 vbRed, User & " has the same or more access than you."
     End If
     
    Else
 
    AccessList.Add LCase(User), AccessLevel  'add
    Chat.AddChat2 vbRed, "Added " & User & " with access " & AccessLevel
    
    End If

End Function
Public Function DoWork(User As String, CMDData As String)

Dim SplitCMD() As String
    SplitCMD = Split(CMDData, " ")

Select Case LCase(SplitCMD(0))  'AccessList.Item(LCase(User))
''''''''Access Level of 10'''''''''
'''''''''''''''''''''''''''''''''''
Case "ver", "about", "version"
    If AccessList.Item(LCase(User)) >= 10 Or LCase(User) = LCase(strMast) Then
      Queue.Enqueue "Game Announcer v" & App.Major & "." & App.Minor & " (Build: " & App.Revision & ") by Myst and cHip."
    End If

Case "find", "whois"
    If AccessList.Item(LCase(User)) >= 10 Or LCase(User) = LCase(strMast) Then
    If CheckAccess(SplitCMD(1)) = True Then
        If LCase(SplitCMD(1)) = LCase(strMast) Then
        Queue.Enqueue SplitCMD(1) & " is the bot's master."
        Else
        Queue.Enqueue SplitCMD(1) & " has " & AccessList.Item(LCase(SplitCMD(1))) & " access"
        End If
      
      Else
      Queue.Enqueue SplitCMD(1) & " was not found."
    End If
    End If
    
Case "say"
    If AccessList.Item(LCase(User)) >= 10 Or LCase(User) = LCase(strMast) Then
       SplitCMD = Split(CMDData, " ", 2)
       If Mid$(SplitCMD(1), 1, 1) = "/" Then
        SplitCMD(1) = Replace(SplitCMD(1), "/", "", 1, 1)
       End If
    Queue.Enqueue SplitCMD(1)
    End If

Case "trigger", "trig", "t"
    If AccessList.Item(LCase(User)) >= 10 Or LCase(User) = LCase(strMast) Then
        Queue.Enqueue "/w " & User & " Current Trigger: " & IniTrigger
    End If

Case "timefilter", "time"
    If AccessList.Item(LCase(User)) >= 70 Or LCase(User) = LCase(strMast) Then
        If UBound(SplitCMD) = 0 Then  'blank time
            Send0x0E "You must input how many seconds you wanted filtered!"
            Exit Function
        End If
        If SplitCMD(1) = "" Then 'user forgot something
             Send0x0E "You must input how many seconds you wanted filtered!"
             Exit Function
        End If
            TimeFilter = Trim(SplitCMD(1))
            Send0x0E "Detecting games within " & TimeFilter & " secs."
    End If

Case "filtercount"
    If AccessList.Item(LCase(User)) >= 20 Or LCase(User) = LCase(strMast) Then
        Send0x0E "You have " & GameFilters2.count & " Filters. {" & Gamefilter.count & " Hashs)"
    End If
    
Case "track"
    If AccessList.Item(LCase(User)) >= 70 Or LCase(User) = LCase(strMast) Then
            HandleTrack Mid$(CMDData, 7)
    End If
    
Case "clearall", "clear"
    If AccessList.Item(LCase(User)) >= 70 Or LCase(User) = LCase(strMast) Then
        Call ClearAll
    End If
 
Case "ep", "tv", "tvshow"
    If AccessList.Item(LCase(User)) >= 10 Or LCase(User) = LCase(strMast) Then
        GotShowID = False
 frmMain.Inet1.Execute "http://services.tvrage.com/feeds/search.php?show=" & Trim(LCase(Mid$(CMDData, Len(SplitCMD(0)) + 2)))  'Smallville"

    End If
 
Case "deletefilter", "delete", "del"
    If AccessList.Item(LCase(User)) >= 70 Or LCase(User) = LCase(strMast) Then
        Call DeleteFilter(Trim(LCase(Mid$(CMDData, Len(SplitCMD(0)) + 2))))
    End If
            
            
Case "switch", "watch"
    If AccessList.Item(LCase(User)) >= 70 Or LCase(User) = LCase(strMast) Then
                Select Case UCase(Trim(SplitCMD(1)))
                Case "UMS"
                WatchType = "UMS"
                Case "TVB"
                WatchType = "TVB"
                Case "MELEE"
                WatchType = "MELEE"
                Case Else
                Send0x0E "Improper Watch Type, use : UMS, TVB, or MELEE"
                Exit Function
                End Select
                    Send0x0E "Switched to detecting " & WatchType & " games."
   End If
         
Case "stop", "sleep", "stoptrack"
    If AccessList.Item(LCase(User)) >= 70 Or LCase(User) = LCase(strMast) Then
        If TimerOn = True Then
         Call KillTimer(frmMain.hWnd, 1000)
            Queue.Enqueue "Stopped Tracking!"
            TimerOn = False
         Else
            Queue.Enqueue "Tracking is already off"
        End If
    End If
            
Case "start", "starttrack"
    If AccessList.Item(LCase(User)) >= 70 Or LCase(User) = LCase(strMast) Then
         If TimerOn = False Then
                    Call SetTimer(frmMain.hWnd, 1000, 10000, AddressOf TimerProc)
                    Queue.Enqueue "Tracking Started"
         Else
                    Queue.Enqueue "Tracking is already on"
         End If
    End If
Case "weather"
    If AccessList.Item(LCase(User)) >= 70 Or LCase(User) = LCase(strMast) Then
        SplitCMD = Split(CMDData, " ", 2)
        HandleWeather SplitCMD(1)
    End If
Case "rem", "remove"
    If AccessList.Item(LCase(User)) >= 80 Or LCase(User) = LCase(strMast) Then
        If LCase(SplitCMD(1)) <> LCase(strMast) Then
            RemoveAccess User, SplitCMD(1)
        Else
            Queue.Enqueue "You can't remove the master's access."
        End If
    End If
    
    
Case "add"
    If AccessList.Item(LCase(User)) >= 80 Or LCase(User) = LCase(strMast) Then
     If LCase(SplitCMD(1)) <> LCase(strMast) Then
        AddAccess User, SplitCMD(1), SplitCMD(2)
        Else
        Queue.Enqueue "You can't change the Master's access."
     End If
    End If

Case "join"
 If AccessList.Item(LCase(User)) >= 70 Or LCase(User) = LCase(strMast) Then
        Queue.Enqueue "/join " & Trim(LCase(Mid$(CMDData, Len(SplitCMD(0)) + 2)))
 End If

Case "language"
 If AccessList.Item(LCase(User)) >= 70 Or LCase(User) = LCase(strMast) Then
        Queue.Enqueue "Changed Language to " & Trim(LCase(Mid$(CMDData, Len(SplitCMD(0)) + 2)))
 End If
 
Case "games", "last", "game"
    If AccessList.Item(LCase(User)) >= 10 Or LCase(User) = LCase(strMast) Then
        Queue.Enqueue Last5Games
    End If
    
Case "listfilters", "filters", "filterlist"
    If AccessList.Item(LCase(User)) >= 20 Or LCase(User) = LCase(strMast) Then
        Queue.Enqueue ListFilters
    End If
    
    
Case "harvest"
    If LCase(SplitCMD(1)) = "on" Then
        harvestbool = True
        Queue.Enqueue "Harvesting Key Hashes = On"
            Else
        harvestbool = False
        Queue.Enqueue "Harvesting Key Hashes = Off"
    End If


 '''''''WINAMP COMMANDS''''''''''''''''''
  Case "mp3", "music"
      If AccessList.Item(LCase(User)) >= 100 Or LCase(User) = LCase(strMast) Then

            WindowTitle = GetWindowTitle("Winamp v1.x")
            If WindowTitle = vbNullString Then
                Send0x0E "Winamp is not active"
            Else
            WindowTitle = Left(WindowTitle, Len(WindowTitle) - 9)
                Send0x0E "Track [" & WindowTitle & "] ..::Game Announcer::.."
                
            End If
        End If
   Case "play", "p"
         If AccessList.Item(LCase(User)) >= 100 Or LCase(User) = LCase(strMast) Then

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
          End If
                Case "next", "n", "skip"
                
                  If AccessList.Item(LCase(User)) >= 100 Or LCase(User) = LCase(strMast) Then

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
          End If
            Case "prev", "back", "last"
                  If AccessList.Item(LCase(User)) >= 100 Or LCase(User) = LCase(strMast) Then

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
            End If
               Case "pause"
                     If AccessList.Item(LCase(User)) >= 100 Or LCase(User) = LCase(strMast) Then

            iwinamp = WinAMP_SendCommandMessage(WA_PAUSE)
            If i = 0 Then
               ' Send0x0E "Winamp Paused"
                
            Else
                Send0x0E "Winamp is not active"
            End If
            End If
    Case "winamp"
If AccessList.Item(LCase(User)) >= 100 Or LCase(User) = LCase(strMast) Then
         
            Select Case SplitCMD(1)
            
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
  End If


Case Else 'unknown cmd

End Select

End Function
Public Function AddAccess(User As String, Person2Add As String, AccessLevel As String)
    Dim UserAccess As Integer
    UserAccess = AccessList.Item(LCase(User))

        Select Case CheckAccess(Person2Add)
            
            Case False
                If UserAccess >= AccessLevel Or LCase(User) = LCase(strMast) Then
                     AccessList.Add LCase(Person2Add), AccessLevel
                     Queue.Enqueue "Added " & Person2Add & " with " & AccessList.Item(LCase(Person2Add)) & " access."
                     SaveDatabase Person2Add, AccessLevel
                        Else
                     Queue.Enqueue "You can't add someone more access than you have."

                End If
                
            Case True
             If UserAccess > AccessList.Item(LCase(Person2Add)) Or LCase(User) = LCase(strMast) Then
              AccessList.Remove (LCase(Person2Add))
              AccessList.Add LCase(Person2Add), AccessLevel
                     Queue.Enqueue "Changed " & Person2Add & "'s access to " & AccessList.Item(LCase(Person2Add)) & "."
                     SaveDatabase Person2Add, AccessLevel
             Else
             Queue.Enqueue "Can't change someones access who's greater or equal to yours."
             End If
        End Select
        

End Function
Public Function RemoveAccess(User As String, Person2Rem As String)


    Dim UserAccess As Integer
    UserAccess = AccessList.Item(LCase(User))
        Select Case CheckAccess(Person2Rem)
            
            Case False
               
                     Queue.Enqueue Person2Rem & " was never in the database."

                
                
            Case True
             If UserAccess > AccessList.Item(LCase(Person2Rem)) Or LCase(User) = LCase(strMast) Then
              AccessList.Remove (LCase(Person2Rem))
                            Queue.Enqueue "Removed " & Person2Rem & " from database."
                     SaveDatabase Person2Rem, AccessLevel, True
             Else
             Queue.Enqueue "Can't remove someone who has more access than you."
             End If
        End Select

End Function
Public Function SaveDatabase(User As String, ByVal AccessLevel As String, Optional IncomingRemove As Boolean)
    Dim intFreeFile As Integer, i As Integer, o As Integer
     i = 1
    Const DELIMETER As String = ";"
     
    intFreeFile = FreeFile
    
    Open (App.Path & "\database.txt") For Output As #intFreeFile
 
 Do Until AccessList.GrabKeyBasedOnINDEX(i) = ""
     
      If AccessList.Exists(AccessList.GrabKeyBasedOnINDEX(i)) = True Then     'ignore its empty slot
        Print #intFreeFile, AccessList.GrabKeyBasedOnINDEX(i) & DELIMETER & AccessList.Item(AccessList.GrabKeyBasedOnINDEX(i))
      End If
      
      i = i + 1
DoEvents
Loop

    Chat.AddChat vbWhite, "Database.txt saved"
   
    Close #intFreeFile
End Function
Public Function CheckAccess(User As String) As Boolean

If AccessList.Item(LCase(User)) <> "" Or LCase(User) = LCase(strMast) Then  'no access
CheckAccess = True
Else
CheckAccess = False
End If

End Function

Public Function LoadAccess(FiileName As String)
Dim whole_file As String
Dim animals() As String
Dim usersNaccess() As String
Dim i As Integer
 
'On Error GoTo checkthis
    ' Get the whole file.
    whole_file = GrabFile(App.Path & "\" & FiileName)
    ' Break the file into lines.
    animals = Split(whole_file, vbCrLf)
     
    Do Until i = UBound(animals) + 1
    
'skip blank lines
If animals(i) = "" Then
Else
''
If InStr(1, animals(i), ";") = 0 Then
MsgBox "Invalid Database.txt format [" & animals(i) & "]. Input your users as Username;Accesslevel"
KillProcess.KillProcess ("GameAnnouncer.exe")
Exit Function
End If
End If
usersNaccess = Split(Mid$(animals(i), 1), ";")
 

    If Mid$(Trim(animals(i)), 1) <> "" Then
If AccessList.Exists(LCase(usersNaccess(0))) = True Then
  
Else
 AccessList.Add LCase(usersNaccess(0)), usersNaccess(1)
  keysz = LCase(usersNaccess(0))
 End If
    End If
   i = i + 1
   DoEvents
   Loop
 
'checkthis: MsgBox ("There is something wrong with your Database.txt, please check it (Format: Name;AccessLevel)")
End Function

   Private Function GrabFile(ByVal file_name As String) As _
    String
Dim fnum As Integer

    On Error GoTo NoFile
    fnum = FreeFile
    Open file_name For Input As fnum
    GrabFile = Input$(LOF(fnum), fnum)
    Close fnum
    Exit Function

NoFile:
    GrabFile = ""
    MsgBox "Couldnt Grab File :( something's wrong!!"
End Function
Public Sub HandleTrack(MSG As String)
 Dim Filter As String
 Dim What2Filter As String
 
 Filter = Trim(LCase((Mid$(MSG, 1, InStr(1, MSG, " ")))))
 What2Filter = Trim(LCase(Mid$(MSG, Len(Filter) + 1)))
 
 
 
 Select Case LCase(Filter)
    Case "map"
    isMapfilter = True
    Case "host"
    isHostfilter = True
    Case "game"
    isGamefilter = True
    Case Else
    Queue.Enqueue "Invalid Filter, proper filters are: map, host, or game."
    Exit Sub
 End Select
 
 If What2Filter = "" Then
 
 Select Case LCase(Filter)
    Case "map"
     Queue.Enqueue "You need to tell me what map to look for."
    Case "host"
     Queue.Enqueue "You need to tell me what host to look for."
    Case "game"
     Queue.Enqueue "You need to tell me what game name to look for."
 End Select
 
 Exit Sub
 Else
 Call AddFilter(Filter, What2Filter)
 
 End If
 
 If FirstTime = True Then
     Call SetTimer(frmMain.hWnd, 1000, 10000, AddressOf TimerProc)
End If
 
End Sub
Public Function AddFilter(FilType As String, Filter As String)

If Gamefilter.Exists(Gamefilter.HashCode("MYST" & Filter)) = True Then
Queue.Enqueue "Already Tracking: " & Filter
 
Else
 Gamefilter.Add Gamefilter.HashCode("MYST" & Filter), Filter
 GameFilters2.Add Filter
 
Queue.Enqueue "Tracking " & FilType & ": " & Filter
 
 
 End If


End Function
Public Function ListFilters() As String

Dim MsgtoDisplay As String
Dim Y As Integer

For Y = 1 To GameFilters2.count

If GameFilters2.count = 0 Then
'do nothing
Else
If MsgtoDisplay = "" Then
MsgtoDisplay = GameFilters2.Item(Y)
Else
MsgtoDisplay = MsgtoDisplay & ", " & GameFilters2.Item(Y)
End If
End If



Next Y
'[12:48:36 PM] - : Announcer : -  Filter(s): 3v3, bgh, micro, mouse, bald locks, rpg, diplo, temple siege
If MsgtoDisplay = "" Then
MsgtoDisplay = "I have no filters added yet."
End If


If Len(MsgtoDisplay) > 210 Then  'count takeing into consderation the "filter(s)" word
Dim HoleMe() As String, splitmsg As String
HoleMe = SplitString(MsgtoDisplay, 210)
Dim w As Integer, i As Integer

For i = 0 To UBound(HoleMe)
Queue.Enqueue "Filter(s): " & HoleMe(i)

Next i

Else
ListFilters = "Filter(s): " & MsgtoDisplay
End If
End Function
Public Function SplitString(ByVal TheString As String, ByVal StringLen As Integer) As String()
Dim ArrCount As Integer
Dim i As Long
Dim TempArray() As String
  ReDim TempArray((Len(TheString) - 1) \ StringLen)
  For i = 1 To Len(TheString) Step StringLen 'icrement it by teh len you wanted it by
    TempArray(ArrCount) = Mid$(TheString, i, StringLen)
    ArrCount = ArrCount + 1
  Next
  SplitString = TempArray
End Function
Public Function DeleteFilter(Filter As String)
 
If Gamefilter.Exists(Gamefilter.HashCode("MYST" & Filter)) = True Then

Gamefilter.Remove Gamefilter.HashCode("MYST" & Filter) 'remove from hashtable
Dim i As Integer
i = 1

Do Until i > GameFilters2.count
    If GameFilters2.Item(i) = Filter Then
    GameFilters2.Remove (i)
     End If
i = i + 1
DoEvents
Loop
 
 Send0x0E "Deleted " & Filter
Else

  
 
Queue.Enqueue Filter & " was never added."
 
 
End If


End Function
Public Function ClearAll()
Gamefilter.RemoveAll
 

Do Until GameFilters2.count = 0
GameFilters2.Remove (1) ' collections drop slots automaticly

 
DoEvents
Loop

isMapfilter = False
isHostfilter = False
isGamefilter = False
Send0x0E "Cleared all Filters"
End Function
Public Function ClearedMapData(MapData As String) As String 'Remove < 0x20 chars from data

Dim M As Long
M = 1

Do Until M >= Len(MapData)
 
 
If buf2.GetBYTE(Mid$(MapData, M, 1)) < 20 Then

MapData = Replace(MapData, Mid$(MapData, M, 1), "")
End If

M = M + 1


DoEvents
Loop

ClearedMapData = Trim(MapData)

 End Function

Public Function CheckFilterMatch(GameData As String, MapData As String, HostData As String) As Boolean
Dim identif As String
Dim cnt As Long
cnt = 1

If GameFilters2.count = 0 Then
Exit Function
End If

Do Until cnt = GameFilters2.count + 1


identif = GameFilters2.Item(cnt)

If InStr(1, LCase(GameData), identif) <> 0 Then
CheckFilterMatch = True
Exit Function
End If



If InStr(1, LCase(MapData), identif) <> 0 Then
CheckFilterMatch = True

Exit Function
End If

If InStr(1, LCase(HostData), identif) <> 0 Then
CheckFilterMatch = True
Exit Function
End If



cnt = cnt + 1
DoEvents
Loop

End Function

Public Function Last5Games() As String

Dim MsgtoDisplay As String
Dim Y As Integer

For Y = 0 To UBound(Last5Gamess)

If Last5Gamess(Y).GameNames = "" Then
'do nothing
Else
If MsgtoDisplay = "" Then
MsgtoDisplay = Last5Gamess(Y).GameNames
Else
MsgtoDisplay = MsgtoDisplay & ", " & Last5Gamess(Y).GameNames
End If
End If



Next Y

If MsgtoDisplay = "" Then
MsgtoDisplay = "No games detected since I started."
End If

Last5Games = MsgtoDisplay

End Function

Public Function Stats(GameName As String, Map As String, Host As String, Time As String)

End Function

Private Sub HandleWeather(Params As String)
    Dim xml_Google As New DOMDocument
    Dim xml_node As IXMLDOMNode
    Dim inet As String, Temp As String, Cond As String, City As String
    Dim Wind As String, Humid As String

    inet = frmMain.Inet1.OpenURL("http://www.google.com/ig/api?weather=" & Params)
    xml_Google.loadXML inet 'loads into the inet
    
    Set xml_node = xml_Google.documentElement.selectSingleNode("//xml_api_reply/weather/forecast_information/city")
        City = xml_node.Attributes.getNamedItem("data").Text
            If (City = vbNullString) = True Then
                Send0x0E "U.S. City does not exist or wrong syntax"
                Exit Sub
            End If
    Set xml_node = xml_Google.documentElement.selectSingleNode("//xml_api_reply/weather/current_conditions/temp_f")
        Temp = xml_node.Attributes.getNamedItem("data").Text
    Set xml_node = xml_Google.documentElement.selectSingleNode("//xml_api_reply/weather/current_conditions/condition")
        Cond = xml_node.Attributes.getNamedItem("data").Text
    Set xml_node = xml_Google.documentElement.selectSingleNode("//xml_api_reply/weather/current_conditions/wind_condition")
        Wind = xml_node.Attributes.getNamedItem("data").Text
    Set xml_node = xml_Google.documentElement.selectSingleNode("//xml_api_reply/weather/current_conditions/humidity")
        Humid = xml_node.Attributes.getNamedItem("data").Text

    Send0x0E City & " - Temperature: " & Temp & "ºF - " & "Currently: " & Cond & " - " & Wind & " - " & Humid & "."
    
    Set xml_node = Nothing
    Exit Sub

End Sub
Public Function ConnectingID(Cdkey As String) As String


Dim Product As Long
Dim PublicValue As Long
Dim PrivateValue As String

Select Case Len(Cdkey)
    Case 26
        If kd_quick(Cdkey, 0, 0, PublicValue, Product, outhash, 20) = 0 Then
        End If
    Case 13, 16
        If decode_hash_cdkey(Cdkey, 0, 0, PublicValue, Product, outhash) = 0 Then
        End If
End Select



 
Select Case Hex(Product)

    Case 1, 2, 17 'SC
    ConnectingID = "SC"
    WhatConnectID = "SC"
    Case 6, 7, "A", "C", 18, 19 'D2
    ConnectingID = "D2"
    WhatConnectID = "D2"
    Case "E", "F", 12, 13 'W3
    ConnectingID = "W3"
    WhatConnectID = "W3"
    Case Else
    ConnectingID = "Unknown"
    

End Select


 

End Function

Public Sub ChangeBackground()
     Select Case modFunctions.ConnectingID(IniCDKey)
     
     Case "SC"
     frmMain.Picture = LoadPicture(App.Path & "\SCBackground.jpg")
     Case "D2"
     frmMain.Picture = LoadPicture(App.Path & "\D2Background.jpg")
     Case "W3"
     frmMain.Picture = LoadPicture(App.Path & "\WCBackground.jpg")
     Case Else
     frmMain.Picture = LoadPicture(App.Path & "\SCBackground.jpg")
     
     End Select
   
   frmMain.txtChan.Refresh
End Sub

Public Function GoogleTranslate(Message As String, Optional OutputLanguage As String = "en") As String

Dim var1 As String 'Holds data that you're POSTing
 
If Len(Message) > 222 Then
GoogleTranslate = "Translation is too long for Battle.net server"
Exit Function
End If

TransON = True
On Error Resume Next
var1 = "q=" & Message & "?&v=1.0&langpair=en%7C" & OutputLanguage
 frmMain.Inet1.Execute "http://ajax.googleapis.com/ajax/services/language/translate", "POST", var1, "Content-Type: application/x-www-form-urlencoded"

End Function

