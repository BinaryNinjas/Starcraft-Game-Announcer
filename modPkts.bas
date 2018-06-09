Attribute VB_Name = "modPkts"
''''This is where all your packets will go
'''We'll need to make the following BNCS packets
''' 0x50, 0x51
''' 0x3A, 0x14
''' 0x0A, 0X0B
''' 0X0C, 0x09
Public Declare Function kd_quick Lib "bncsutil.dll" _
    (ByVal CDKey As String, ByVal ClientToken As Long, _
    ByVal ServerToken As Long, PublicValue As Long, _
    Product As Long, ByVal HashBuffer As String, _
    ByVal BufferLen As Long) As Long

Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Declare Function SetTimer Lib "User32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "User32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public TimerOn As Boolean
Public WatchType As String
Public buf2 As New clsBuffer 'Publically declare buf2 to use functions in class
Public UDPToken As Long 'Grabbed from S>C 0x50
Public ServerToken As Long  'Grabbed from S>C 0x50
Public tmpFileName As String 'Grabbed from S>C 0x50
Public tmpFormula As String 'Grabbed from S>C 0x50
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Sub CreateGame_0x1C(Gamename As String, GamePass As String) ', gamestatstring As String, MapName As String, GameSpeed As String, MapType As String, MapIcon As String, MapSize As Integer, GameType As String)
Chat.AddChat2 vbBlack, "Creating Game"
Dim count As Long
count = GetTickCount
With buf2
.InsertDWORD &H0 'game state
.InsertDWORD &H0 'Time since creation
'.InsertWORD &H0 'game type
.InsertDWORD &HA '"&H" & GameType  '&H2  'Parameter
.InsertWORD &H1  'Unknown
.InsertWORD &HFF 'Unknown
.InsertDWORD &H0 'ladder is no
.InsertNTString Gamename
.InsertNTString GamePass
'Game Stat String Below
.InsertSTRING ",,,6,6,a,,1,34eab02f,4,," 'Game Options
'.InsertSTRING "," & MapSize & ",," & GameSpeed & "," & MapIcon & "," & GameType & ",,1" & ",34eab02f," & MapType & ",,"
'44 = map size
'6 = game speed (Fastest)
'1 = map icon
'2 = game type (melee)
'1 = game penalty 1 is no penalty on disc
'after that is timestamp
.InsertNT_0DString "My$t"  '"Thief" 'Username
.InsertNT_0DString "www.DarkBlizz.org" '"The Hunters" 'MapName
.InsertBYTE &H0
.InsertHEADER &H1C
.SendPacket frmMain.sckBnet
End With


End Sub
 

Public Sub Send0x09() '0x09 UDP
Call Chat.AddChat(vbWhite, "Sending UDP Token")
Dim P As Integer

For P = 1 To 3
With buf2
 .InsertDWORD &H9
.InsertDWORD ServerToken
.InsertDWORD UDPToken
.SendPacketUDP frmMain.sckUDP
End With
Next P

End Sub
Public Sub Send0x09TCP()
'0030  ff a5 2e 18 00 00 ff 09  17 00 0a 00 00 00 ff ff   ........ ........
'0040  00 00 00 00 00 00 19 00  00 00 00 00 00            ........ .....
'Chat.AddChat vbWhite, "Sending 0x09TCP"

With buf2
.InsertWORD &HA 'shows melee
.InsertWORD &H0 ' condition 2
.InsertBYTE &HFF 'condiiton 3
.InsertBYTE &HFF
.InsertWORD &H0
.InsertDWORD &H0 'condition 4
.InsertDWORD &H19
.InsertBYTE &H0
.InsertBYTE &H0
.InsertBYTE &H0
.InsertHEADER &H9
.SendPacket frmMain.sckBnet
End With

End Sub
Public Sub Send0x09TCP2()
'0030  fb 09 61 05 00 00 ff 09  17 00 0a 00 00 00 ff ff   ..a..... ........
'0040  00 00 00 00 00 00 19 00  00 00 00 00 00            ........ .....
'Chat.AddChat vbWhite, "Sending 0x09TCP"

With buf2
.InsertWORD &HA
.InsertWORD &H0
.InsertDWORD &H0
.InsertWORD &HFF
.InsertWORD &HFF
.InsertDWORD &H1
.InsertNTString "gName"
.InsertNTString ""
.InsertNTString ""
.InsertHEADER &H9
.SendPacket frmMain.sckBnet
End With

End Sub
Public Sub Send0x09TCPtvb()
'0030  ff a5 2e 18 00 00 ff 09  17 00 0a 00 00 00 ff ff   ........ ........
'0040  00 00 00 00 00 00 19 00  00 00 00 00 00            ........ .....
'Chat.AddChat vbWhite, "Sending 0x09TCP"

With buf2
.InsertWORD &HF  'shows melee
.InsertWORD &H0 ' condition 2
.InsertBYTE &HFF 'condiiton 3
.InsertBYTE &HFF
.InsertWORD &H0
.InsertDWORD &H0 'condition 4
.InsertDWORD &H19
.InsertBYTE &H0
.InsertBYTE &H0
.InsertBYTE &H0
.InsertHEADER &H9
.SendPacket frmMain.sckBnet
End With

End Sub
Public Sub Send0x09TCPmelee()
'0030  ff a5 2e 18 00 00 ff 09  17 00 0a 00 00 00 ff ff   ........ ........
'0040  00 00 00 00 00 00 19 00  00 00 00 00 00            ........ .....
'Chat.AddChat vbWhite, "Sending 0x09TCP"

With buf2
.InsertWORD &H2 'shows melee
.InsertWORD &H0 ' condition 2
.InsertBYTE &HFF 'condiiton 3
.InsertBYTE &HFF
.InsertWORD &H0
.InsertDWORD &H0 'condition 4
.InsertDWORD &H19
.InsertBYTE &H0
.InsertBYTE &H0
.InsertBYTE &H0
.InsertHEADER &H9
.SendPacket frmMain.sckBnet
End With

End Sub
Public Sub Send0x50(verbyte As String)
Call Chat.AddChat(vbGreen, "Sending 0x50")

    With buf2
        .InsertDWORD &H0
        .InsertSTRING "68XI"   '&H49583836 '68XI"
        .InsertDWORD .GetDWORD("PXES")   'turn string into number into packer buffer
        .InsertDWORD "&H" & verbyte   'version byte for sc/bw
        .InsertDWORD &H0    ' All
        .InsertDWORD &H0    ' This
        .InsertDWORD &H0    ' Can
        .InsertDWORD &H0    ' Be
        .InsertDWORD &H0    ' Null
        .InsertNTString "GBR" 'null byte terminated string
        .InsertNTString "United Kingdom"
        .InsertHEADER &H50 'add the header FF 50 and 2 byte lengh of packet buffer
        .SendPacket frmMain.sckBnet   'hand the winsock to the sub so it can use it to send the data down it

    End With
    
End Sub

Public Sub Send0x51(tmpFormula As String, tmpFileName As String, Checksum As Long, ByVal exeversion As Long, EXEInfo As String)
Call Chat.AddChat(vbGreen, "Sending 0x51")
Dim Product As Long
Dim PublicValue As Long
Dim PrivateValue As String
Dim CDKey As String

CDKey = IniCDKey

 
    
' (DWORD) Client Token
' (DWORD) EXE Version
'(DWORD) EXE Hash
'(DWORD) Number of CD-keys in this packet
'(BOOLEAN) Spawn CD-key **

'For Each Key:

'    (DWORD) Key Length
'    (DWORD) CD-key's product value
'    (DWORD) CD-key's public value
'    (DWORD) Unknown (0)
'    (DWORD) [5] Hashed Key Data


'(STRING) Exe Information
'(STRING) CD-Key owner name *

    Dim ClientToken As Long
    ClientToken = GetTickCount 'any old number (cant be 0)
    

 'We need to remove the extra nullbytes from the extra allocated space
  EXEInfo = Mid$(EXEInfo, 1, InStr(1, EXEInfo, vbNullChar) - 1) 'remove extra null bytes
 ''''
 
   Dim outhash As String
   outhash = String(20, 0) 'allocate space
 
Select Case Len(CDKey)

Case 26
If kd_quick(CDKey, ClientToken, ServerToken, PublicValue, Product, outhash, 20) = 0 Then
Chat.AddChat vbRed, "Failed to hash key, make sure you're using a valid Starcraft CDKey!"
frmMain.sckBnet.Close
Exit Sub
End If
 
Case 13
If decode_hash_cdkey(CDKey, ClientToken, ServerToken, PublicValue, Product, outhash) = 0 Then
Chat.AddChat vbRed, "Failed to hash key, make sure you're using a valid Starcraft CDKey!"
frmMain.sckBnet.Close
Exit Sub
End If

End Select

' Call calcHashBuf(KeyHash, Len(KeyHash), outhash)

 With buf2
.InsertDWORD ClientToken 'GetTickCount()
 .InsertDWORD exeversion
 .InsertDWORD Checksum    'EXE hash
.InsertDWORD &H1 '1 cdkey
 .InsertDWORD &H0 'no spawning
 .InsertDWORD Len(CDKey) 'length of key
.InsertDWORD Product  'Product Value of Key i.e 01 or 02
.InsertDWORD PublicValue  'Public Value of Cd Key 7digit number
 .InsertDWORD &H0 'Null
.InsertSTRING outhash 'Hashed Key Data
.InsertNTString EXEInfo   'Exe Info
 .InsertNTString "GameAnnouncer" 'Owner Name
.InsertHEADER &H51
 .SendPacket frmMain.sckBnet
 
 End With

End Sub

Public Sub Send0x3A()
'(DWORD) Client Token
'(DWORD) Server Token
'(DWORD) [5] Password Hash
'(STRING) Username
Chat.AddChat vbGreen, "Sending 0x3A"
Dim outhash As String
    outhash = String(20, 0)
ClientToken = GetTickCount

Call double_hash_password(LCase(IniPass), ClientToken, modPkts.ServerToken, outhash)

   
    With buf2
        .InsertDWORD ClientToken
        .InsertDWORD modPkts.ServerToken
        .InsertSTRING outhash
        .InsertNTString IniName
        .InsertHEADER &H3A
        .SendPacket frmMain.sckBnet
    End With
        
End Sub
Public Sub Send0x3D()
  Chat.AddChat vbGreen, "Sending 0x3D"
    Dim outhash As String * 20
    'Call Chat.ShowChat(i, vbGreen, "Sending 0x3D")
outhash = String(20, 0)
 
Call hash_password(LCase(IniPass), outhash)

With buf2
.InsertSTRING outhash
.InsertNTString IniName
.InsertHEADER &H3D
.SendPacket frmMain.sckBnet
End With



End Sub
Public Sub Send0x14()
'MsgBox udpstamp
Chat.AddChat vbGreen, "Sending 0x14"
With buf2
 '.InsertDWORD udpstamp
 .InsertSTRING "tenb"
 .InsertHEADER &H14
 .SendPacket frmMain.sckBnet
 End With
End Sub

Public Sub Send0x0A()
Chat.AddChat vbGreen, "Sending 0x0A"


    With buf2
        .InsertNTString IniName
        .InsertNTString "" 'Null
        .InsertHEADER &HA
        .SendPacket frmMain.sckBnet
        
    End With
    
End Sub


Public Sub Send0x0B()

    With buf2
        .InsertSTRING "PXES"
        .InsertHEADER &HB
        .SendPacket frmMain.sckBnet
        
    End With
    
End Sub

Public Sub Send0x0C()
Chat.AddChat vbGreen, "Sending 0x0C"
    With buf2
        .InsertDWORD &H2
        .InsertNTString IniHome
        .InsertHEADER &HC
        .SendPacket frmMain.sckBnet
    End With


End Sub
Public Sub Send0x0E(Text As String)

    With buf2
        .InsertNTString Text
        .InsertHEADER &HE
        .SendPacket frmMain.sckBnet
    End With

Chat.AddChat "&H" & StrReverse("736AFF"), "<" & IniName & "> " & Text
    
End Sub
Public Sub Send0x25(Data As String)

    With buf2
        .InsertDWORD buf2.GetDWORD(Mid(Data, 5, 4))
        .InsertHEADER &H25
        .SendPacket frmMain.sckBnet
    End With
    
End Sub
Public Sub Send0x65()
With buf2

.InsertHEADER &H65
.SendPacket frmMain.sckBnet

End With

End Sub
Public Sub ParseBinary(strData As String)
Dim Flags&, Ping&, Username$, Statstring$
If Len(strData) < 4 Then Exit Sub
    GetValues strData, Ping, Flags, Username, Statstring
        Select Case Asc(Mid(strData, 5, 1))
            Case ID_WHISPFROM: RecvWhisper Username, Flags, Statstring
            Case ID_TALK: Talk Username, Flags, Statstring, Ping
            Case ID_EMOTE: Emote Username, Flags, Statstring
            Case ID_JOIN: Join Username, Flags, Statstring, Ping
            Case ID_USER: User Username, Flags, Statstring, Ping, strData
            'Case ID_FLAGS: Event_Flags Username, Flags, Statstring, Ping
            'Case ID_ERROR: Event_Error Statstring
            Case ID_WHISPTO: SendWhisper Username, Flags, Statstring, Ping
            Case ID_CHAN: Channel Statstring
            Case ID_INFO: Info Statstring
            Case ID_LEAVE: Leave Username, Flags
            Case SID_FRIENDSLIST: Chat.AddChat vbRed, strData
            'Case Else:  Unknown strData
        End Select
End Sub
Public Sub Channel(Statstring As String)
Chat.AddChat vbYellow, "Joined Channel " & Statstring
frmMain.Label1.Caption = IniName & " [" & Statstring & "]"

End Sub
Public Sub Emote(User As String, Flags As Long, Text As String)
On Error Resume Next
If InStr(1, User, "*") Then User = Mid(User, InStr(1, User, "*"))
    If Flags = "17" Then
        Chat.AddChat vbBlue, "<" & User & " " & Text & ">"
    Else
        Chat.AddChat vbYellow, "<" & User & " " & Text & ">"
    End If
End Sub
Private Sub Talk(User As String, Flags, Text As String, Ping)
On Error Resume Next
If InStr(1, User, "*") Then User = Mid(User, InStr(1, User, "*"))
    If Flags = "17" Then
        Chat.AddChat vbWhite, ("<" & User & "> " & Text)
    ElseIf Flags = BNFLAGS_OP Then
        Chat.AddChat vbWhite, ("<" & User & "> " & Text)
    ElseIf Flags = BNFLAGS_MOD2 Then
        Chat.AddChat vbWhite, ("<" & User & "> " & Text)
    Else
        Chat.AddChat3 vbWhite, Text, User
    End If
    If Left$(Text, 1) = IniTrigger Then
        'CommandHandler.ParseCommand User, Mid$(Text, 2, Len(Text) - 1)
        DoWork User, Mid$(Text, 2, Len(Text) - 1)
    ElseIf PrepareCheck(Text) = "?trigger" Or _
        PrepareCheck(Text) = "?trig" Or _
        PrepareCheck(Text) = "?t" Then
        CommandHandler.ParseCommand User, Mid$(Text, 2, Len(Text) - 1)
    End If
End Sub
Private Sub Join(Account As String, Flags, Statstring, Ping)
On Error Resume Next
If InStr(1, Account, "*") Then Account = Mid(Account, InStr(1, Account, "*"))
    Chat.AddChat vbGreen, Account & " has joined the channel. [" & Ping & "ms]"
End Sub
Private Sub Leave(User As String, Flags)
On Error Resume Next
If InStr(1, User, "*") Then User = Mid(User, InStr(1, User, "*"))
        If Flags = "17" Then
            Chat.AddChat vbBlue, (User & " has left the channel")
        Else
            Chat.AddChat vbRed, (User & " has left the channel")
        End If
End Sub
Private Sub User(User As String, Flags, Statstring, Ping, RealStatString)
On Error Resume Next



End Sub
Private Sub SendWhisper(ToUser As String, Flags, Text As String, Ping)
If InStr(1, ToUser, "*") Then ToUser = Mid(ToUser, InStr(1, ToUser, "*"))
    If Not LCase(ToUser) = LCase(Username) Then
        Chat.AddChat "&H" & StrReverse("99FF33"), "<To: " & ToUser & "> " & Text
    End If
End Sub
Private Sub RecvWhisper(User As String, Flags, Text)
    If Flags = "17" Then
        Chat.AddChat "&H" & StrReverse("99FF33"), ("<From: " & User & "> " & Text)
    ElseIf Flags = BNFLAGS_OP Then
        Chat.AddChat "&H" & StrReverse("99FF33"), ("<From: " & User & "> " & Text)
    ElseIf Flags = BNFLAGS_MOD2 Then
        Chat.AddChat "&H" & StrReverse("99FF33"), ("<From: " & User & "> " & Text)
    Else
        Chat.AddChat "&H" & StrReverse("99FF33"), ("<From: " & User & "> " & Text)
    End If
    
    If Left$(Text, 1) = IniTrigger Then
        DoWork User, Mid$(Text, 2, Len(Text) - 1)
    End If
    
End Sub
Private Sub Info(Info As String)
Chat.AddChat vbYellow, Info
End Sub
Function MakeLong(x As String) As Long
    If Len(x) < 4 Then
        Exit Function
    End If
    CopyMemory MakeLong, ByVal x, 4
End Function
Public Sub strcpy(ByRef Source As String, ByVal nText As String)
    Source = Source & nText
End Sub

Public Function KillNull(ByVal Text As String) As String
    Dim i As Integer
    i = InStr(1, Text, Chr(0))
    If i = 0 Then
        KillNull = Text
        Exit Function
    End If
    KillNull = Left(Text, i - 1)
End Function
Public Sub GetValues(ByVal DataBuf As String, ByRef Ping As Long, ByRef Flags As Long, ByRef Name As String, ByRef Txt As String)
    Dim a As Long, B As Long, C As Long, d As Long, e As Long, f As Long, recvbufpos As Long
    Name = ""
    Txt = ""
    recvbufpos = 5
    f = MakeLong(Mid$(DataBuf, recvbufpos, 4))
    recvbufpos = recvbufpos + 4
    a = CVL(Mid$(DataBuf, recvbufpos, 4))
    recvbufpos = recvbufpos + 4
    B = MakeLong(Mid$(DataBuf, recvbufpos, 4))
    recvbufpos = recvbufpos + 4
    C = MakeLong(Mid$(DataBuf, recvbufpos, 4))
    recvbufpos = recvbufpos + 4
    d = MakeLong(Mid$(DataBuf, recvbufpos, 4))
    recvbufpos = recvbufpos + 4
    e = MakeLong(Mid$(DataBuf, recvbufpos, 4))
    recvbufpos = recvbufpos + 4
    Flags = a
    Ping = B
    Call strcpy(Name, KillNull(Mid$(DataBuf, 29)))
    Call strcpy(Txt, KillNull(Mid$(DataBuf, Len(Name) + 30)))
End Sub
Function CVL(x As String) As Long
    If Len(x) < 4 Then
        MsgBox "CVL(): String too short"
        Stop
    End If
    CopyMemory CVL, ByVal x, 4
End Function

Function TimerProc(hWnd As Long, uMsg As Long, EventID As Long, dwTime As Long) As Long
 TimerOn = True
If FirstTime = True Then
Send0x09TCP2
Else
Select Case WatchType
Case "UMS"
Send0x09TCP
Case "TVB"
Send0x09TCPtvb
Case "MELEE"
Send0x09TCPmelee
End Select
End If
FirstTime = False

End Function

