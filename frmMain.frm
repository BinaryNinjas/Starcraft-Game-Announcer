VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTTONS.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "GameAnnouncer - by Myst and cHip"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8970
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":F172
   ScaleHeight     =   5640
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   8760
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox TrayIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   5040
      Picture         =   "frmMain.frx":155B4
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   300
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   255
      Left            =   8520
      TabIndex        =   3
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "X"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   0
      FCOL            =   192
      FCOLO           =   16777215
      EMBOSSM         =   16777088
      EMBOSSS         =   16744703
      MPTR            =   0
      MICON           =   "frmMain.frx":24726
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmdCon 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   40
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Connect"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   0
      FCOL            =   16744576
      FCOLO           =   12648384
      EMBOSSM         =   12632064
      EMBOSSS         =   16777152
      MPTR            =   0
      MICON           =   "frmMain.frx":24742
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin MSWinsockLib.Winsock sckUDP 
      Left            =   1560
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock sckBnet 
      Left            =   1080
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrFlood 
      Interval        =   2750
      Left            =   360
      Top             =   5760
   End
   Begin VB.TextBox txtSend 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   140
      TabIndex        =   0
      Text            =   "Game Announcer ~ Created by Myst & cHip"
      Top             =   5160
      Width           =   8520
   End
   Begin MSWinsockLib.Winsock sckChat 
      Left            =   0
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin LVbuttons.LaVolpeButton cmdDiscon 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   40
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Disconnect"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   0
      FCOL            =   16744576
      FCOLO           =   12648384
      EMBOSSM         =   0
      EMBOSSS         =   0
      MPTR            =   0
      MICON           =   "frmMain.frx":2475E
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   255
      Left            =   8160
      TabIndex        =   4
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   0
      FCOL            =   192
      FCOLO           =   16777215
      EMBOSSM         =   16777088
      EMBOSSS         =   16744703
      MPTR            =   0
      MICON           =   "frmMain.frx":2477A
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   255
      Left            =   7800
      TabIndex        =   7
      ToolTipText     =   "Settings"
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "S"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   0
      FCOL            =   192
      FCOLO           =   16777215
      EMBOSSM         =   16777088
      EMBOSSS         =   16744703
      MPTR            =   0
      MICON           =   "frmMain.frx":24796
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin RichTextLib.RichTextBox txtChan 
      Height          =   3615
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6376
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":247B2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF8080&
      Height          =   300
      Left            =   100
      Top             =   5150
      Width           =   8690
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public timrcounter As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

 Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public UpdateChecker As Boolean
      Private Const GWL_STYLE = (-16)
      Private Const GWL_EXSTYLE = (-20)
      Private Const WS_EX_LAYERED = &H80000
      Private Const LWA_COLORKEY = &H1
        Const WS_EX_TRANSPARENT = &H20&
    Public GameCol As Collection
        Dim CntThis As Integer
        Dim FirstRollover As Boolean
Public WindowProc As Long
Dim udpcount As Integer
Dim udpst As Long 'udpstamp to be passed
Public strLastUser As String
Private Sub cmdCon_Click()
'First close any possible connection that still may be active
   sckBnet.Close
'Then connect to server, with the designated bncs port
Chat.AddChat vbWhite, "Connecting to: " & IniServer
   sckBnet.Connect IniServer, 6112
   
   
End Sub

 

 
 

Private Sub Command1_Click()
 modFunctions.ClearedMapData ("Special FOrces")
 modFunctions.ClearedMapData ("DBZ dsfdsf 9879ds dsf dsf")
   
End Sub

Private Sub Form_Load()
    LoadCFG
    FirstRollover = True
UpdateChecker = True
SetWindowLong txtChan.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT

Me.BackColor = vbCyan 'Set the backcolor of the tobe transparent form to w/e color

          SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED

'now everything is transparent so you're gonna have to use this to only make what is a certain color transparent

          SetLayeredWindowAttributes Me.hWnd, vbCyan, 0&, LWA_COLORKEY
    'AddRTB frmMain.txtChan, clrBotInfo, "Game Announcer v" & App.Major & "." & App.Minor & " (Build: " & App.Revision & ") by Myst and cHip."
   Chat.AddChat2 vbGreen, "Game Announcer v" & App.Major & "." & App.Minor & " (Build: " & App.Revision & ") by Myst and cHip."
   If GrabWebTitle("http://darkblizz.org/thief/GameAnnouncer/UpdateChecker.html") = "v" & App.Major & "." & App.Minor & " Build " & App.Revision Then
  ' Chat.AddChat2 &HFFCCCC, "You are using the lastest version."
   Else
   Chat.AddChat2 &H3333FF, "Your bot is out of date, please visit www.DarkBlizz.org to obtain the latest release."

   End If
    
   'Chat.AddChat2 "&H" & StrReverse("66FFCC"), "www.DarkBlizz.org"
   'Chat.AddChat2 "&H" & StrReverse("CCFF00"), "Op Null [USEast]"
   Chat.AddChat2 "&H" & StrReverse("66FFCC"), "www.DarkBlizz.org"

   Chat.AddChat2 "&H" & StrReverse("CCFF00"), "Op FallenArms ¤ Op Legacy [USEast]"
   Chat.AddChat2 "&H" & StrReverse("6600CC"), "Detecting " & WatchType & " games."
    trigger = ReadINI("Options", "Trigger")
    
    frmMain.Caption = "Game Announcer " & App.Major & "." & App.Minor & " (Build: " & App.Revision & ")"
    Set UserList = New Collection
    Set Queue = New clsQueue
     
    Set CommandHandler = New clsCommandHandler
    'Database.LoadDatabase
    LoadAccess "Database.txt"
    FirstTime = True
    TimeFilter = 30  'deafult 3o secs
Label1.Caption = IniName & " [" & IniHome & "]"
End Sub
Private Sub cmdDiscon_Click()
KillTimer frmMain.hWnd, 1000

    sckBnet.Close
    'AddTimeStamp frmMain.txtChan
    'AddRTB frmMain.txtChan, vbRed, "Disconnected"
    Chat.AddChat vbRed, "Disconnected"
    frmMain.HandleUnload
    frmMain.Caption = "Game Announcer " & App.Major & "." & App.Minor & " (Build: " & App.Revision & ")"
    
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
ReleaseCapture
SendMessage hWnd, WM_NCLBUTTONDOWN, 2, 0&
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape2.BorderWidth = 1
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
Dim ChunkVar As Variant
Dim ChunkStr As String
Select Case State



Case 12 'getchuck
'Chat.AddChat frmGenesis.TabStrip1.SelectedItem.Index, "The request has been completed and all data has been received.", vbWhite
 ChunkVar = Inet1.GetChunk(1024, icString)
ChunkStr = ChunkStr & ChunkVar

 Do
            DoEvents
         ChunkVar = Inet1.GetChunk(1024, icString)
            If Len(ChunkVar) = 0 Then Exit Do
            ChunkStr = ChunkStr & ChunkVar
        Loop
        
 


If GotShowID = False Then
TVRageEpisode_GrabShowID ChunkStr
Else
TVRageEpisode_GrabShowINFO ChunkStr
End If

On Error Resume Next
 

Case Else

'Chat.AddChat frmGenesis.TabStrip1.SelectedItem.Index, "Unknown Inet State: " & Inet4.ResponseCode & " - " & Inet4.ResponseInfo, vbRed

End Select
End Sub
Public Function TVRageEpisode_GrabShowID(XMLData As String)
    Dim XMLArray() As String 'Array to hold split XML Data
    Dim Elements As String 'ID of stuff
    Dim ElementData As String 'Stuff between the < >
    Dim X As Integer 'count looper
        X = 1
        XMLArray = Split(XMLData, Chr(&HA)) 'Split incoming XML data by New Line(0x0A)
  
            Do Until X = UBound(XMLArray) + 1

                If XMLArray(X) <> "" Then 'Skip empty lines
                Elements = Mid$(XMLArray(X), 2, InStr(1, Mid$(XMLArray(X), 2), ">") - 1) 'Parse out each element
                ElementData = Replace(Mid$(XMLArray(X), Len(Elements) + 3, InStr(1, Mid$(XMLArray(X), Len(Elements) + 3), "<")), "<", "")
                End If
 
                    Select Case Elements 'Find the elements

                        Case "Results"
                        If ElementData = "0" Then
                        Send0x0E "Zero Results Found, try to be more specific."
                        'Chat.AddChat Index, "Zero Results Found", vbRed
                        Exit Function
                        End If
                        
                        Case "showid"
                        
                        If Inet1.StillExecuting = True Then
                        pause (2)
                        End If
                         
                        
                        If GotShowID <> True Then
                        Inet1.Execute "http://services.tvrage.com/feeds/episodeinfo.php?sid=" & ElementData
                        End If
                        
                        GotShowID = True
                        
                        Case "name"
                        
                    End Select
            X = X + 1
            DoEvents
            Loop

End Function
Public Function TVRageEpisode_GrabShowINFO(XMLData As String)
    Dim XMLArray() As String 'Array to hold split XML Data
    Dim Elements As String 'ID of stuff
    Dim ElementData As String 'Stuff between the < >
    Dim X As Integer 'count looper
    Dim ShowName, WhenOn, TimeOn, Network, LatestEP, LatestEPDate, LatestEPTitle, NextEP, NextEPDate, NextEPTitle As String
        X = 1
        XMLArray = Split(XMLData, "><") 'Split incoming XML data by ><
  
            Do Until X = UBound(XMLArray) + 1
'MsgBox XMLArray(X)
                If XMLArray(X) <> "" Then 'Skip empty lines
                Elements = Replace(Mid$(XMLArray(X), 1, InStr(1, Mid$(XMLArray(X), 1), ">")), ">", "")    'Parse out each element
                ElementData = Replace(Mid$(XMLArray(X), Len(Elements) + 2, InStr(1, Mid$(XMLArray(X), Len(Elements) + 2), "<")), "<", "")
            
                End If
 
                    Select Case Elements 'Find the elements

                        Case "Results"
                        If ElementData = "0" Then
                        Send0x0E "Zero Results Found, try to be more specific."
                        Exit Function
                        End If
                        
                       
                        Case "name"
                        ShowName = ElementData
                        Case "airtime"
                         WhenOn = ElementData
                         
                         Case "number"
                         If LatestEP = "" Then
                         LatestEP = ElementData
                         Else
                         NextEP = ElementData
                         End If
                         
                         Case "title"
                         If LatestEPTitle = "" Then
                         LatestEPTitle = ElementData
                         Else
                         NextEPTitle = ElementData
                         End If
                         
                         Case "airdate"
                         If LatestEPDate = "" Then
                         LatestEPDate = ElementData
                         Else
                         NextEPDate = ElementData
                         End If
                         
                         
                         
                    End Select
            X = X + 1
            DoEvents
            Loop

'Chat.AddChat Index, ShowName & "--" & WhenOn & " [Latest EP: " & LatestEP & " " & LatestEPTitle & " @" & LatestEPDate & ".] [Next EP " & NextEP & " " & NextEPTitle & " @" & NextEPDate & ".]", vbWhite
'sckIRC(Index).SendData "PRIVMSG " & Chat.Home(Index) & " " & ShowName & "--" & WhenOn & " [Latest EP: " & LatestEP & " " & LatestEPTitle & " @" & LatestEPDate & ".] [Next EP " & NextEP & " " & NextEPTitle & " @" & NextEPDate & ".]" & vbCrLf
Send0x0E ShowName & "--" & WhenOn & " [Latest EP: " & LatestEP & " " & LatestEPTitle & " @" & LatestEPDate & ".] [Next EP " & NextEP & " " & NextEPTitle & " @" & NextEPDate & ".]"
'GotShowID = False
End Function
Sub pause(interval As String)

    Dim X
    X = Timer


    Do While Timer - X < Val(interval)


        DoEvents
        Loop

    End Sub
Private Sub LaVolpeButton1_Click()
 If MsgBox("Are you sure you want to quit?", vbOKCancel) = vbOK Then
    'Database.SaveDatabase
    WriteINI "Options", "Trigger", IniTrigger

    Set CommandHandler = Nothing
   
    Set UserList = Nothing
     Call KillTimer(frmMain.hWnd, 1000)
Unload Me
KillProcess.KillProcess ("GameAnnouncer.exe")

Else
'nothing
End If

End Sub

Private Sub LaVolpeButton2_Click()
NoSysIcon INTRAY
End Sub

Private Sub LaVolpeButton3_Click()
frmSettings.Show
End Sub

Private Sub sckBnet_Connect()
Chat.AddChat vbGreen, "Connected!"
'////i see it, you dont have UDP_Getdata?
'Since we have connected now you send battle.net a 0x01 byte
'telling them we're using the game protocol
'after you send that out you then send 0x50.
'////
'send 0x01
'send 0x50
 

'Bind UDP Socket to BNET Server (Removed 1/4/11)
 'Dim watport As Integer
  ' watport = "6113"
   '  sckUDP.RemoteHost = IniServer
    ' sckUDP.RemotePort = 6112

     ' Do Until FreeBind(sckUDP, watport) = True
      '  watport = watport + 1
      'Loop
       '       Chat.AddChat vbGreen, "UDP Port Binded"
''''''''''''''''''
'MsgBox watport
'what else are udoing?
''what you got multiple router ips uh, what?
buf2.InsertBYTE &H1
buf2.SendPacket sckBnet
Call Send0x50(IniVerByte)

End Sub
Public Function FreeBind(SOCK As Winsock, ByVal PortRange As Long) As Boolean
    On Error GoTo NextPORT
    With SOCK
        .Close
        .Bind PortRange, .LocalIP
    End With
    FreeBind = True
    Exit Function
NextPORT:
    If Err.Number = 10048 Then
        FreeBind = False
    Else
        FreeBind = True
        Chat.AddChat vbRed, "[UDP] Bind Error:" & Err.Number & " Unable to bind " & Err.Description
    End If
End Function

Private Sub sckBnet_DataArrival(ByVal bytesTotal As Long)
Static RecvBuffer As String 'we hold the data sent to us from bnet in this string
    Dim PacketLengh As Integer
    Dim Data As String
    Call sckBnet.GetData(Data, vbString)
    RecvBuffer = RecvBuffer & Data 'add the recv data to are buffer string from bnet
    'now lets split each packet up and parse it baced on the 2 bytes lengh in the header
    
    
    While Len(RecvBuffer) >= 4
        PacketLengh = buf2.GetWORD(Mid(RecvBuffer, 3, 2))
        If Len(RecvBuffer) < PacketLengh Then Exit Sub 'we have to wait for the missing data to come before we handle it (packet may have got split up abit)
        'cut out the packet from the buffer and put it in are tmp data string
        Data = Left(RecvBuffer, PacketLengh)
        RecvBuffer = Mid(RecvBuffer, 1 + PacketLengh) 'skip the packet we just cut out
        Call ParseBnetData(Data)     'see a few subs below where we hand the packet to the parse sub
    Wend
    

End Sub
Private Sub ParseBnetData(ByVal Data As String)

Dim PacketID As Byte 'Holds unsigned 8-bit (1-byte) integers ranging in
                     'value from 0 through 255.
                     'http://msdn.microsoft.com/en-us/library/e2ayt412%28VS.80%29.aspx
                     
                     
PacketID = Asc(Mid(Data, 2, 1)) 'Returns the value from the 2nd byte, the packet id, from incoming data
  
'''Now we need to parse out all these Packet ID's
'''We'll use the Select Case Statement, due to the multiple different statements that can occur
'''http://msdn.microsoft.com/en-us/library/cy37t14y.aspx


    Select Case PacketID
Case &H1C 'GameCreateOK or Not OK
  
  If buf2.GetDWORD(Mid(Data, 5, 4)) = &H0 Then  'good
  Call Chat.AddChat(vbRed, "Game Created Successfully")
  Send0x0E "Game created Successfully"
  Else
  Call Chat.AddChat(vbRed, "Game Failed to Create")
  Send0x0E "Game Failed to Create!"
  End If
  
  
Case &H50   'We recieve packet 0x50 (http://www.bnetdocs.org/?op=packet&pid=146)
''' From the packet log below you see the BNCS identifier FF, and then next to that the Packet ID 50.
''' Following that you see 3E 00 (WORD).  That is the length of the entire packet, including the
''' FF 50 in the begining.

'0030  ff c4 42 fc 00 00 ff 50  3e 00 00 00 00 00 e7 9a   ..B....P >.......
'0040  6d 51 09 29 37 00 00 8b  51 03 70 5f c7 01 6c 6f   mQ.)7... Q.p_..lo
'0050  63 6b 64 6f 77 6e 2d 49  58 38 36 2d 31 38 2e 6d   ckdown-I X86-18.m
'0060  70 71 00 b4 ce 3d dc 82  53 16 61 3e 05 23 5a 3b   pq...=.. S.a>.#Z;
'0070  bb e5 92 00                                        ....

''''// Packet Header: ff 50 3e 00
''''//(DWORD) Logon Type: 00 00 00 00
''''//(DWORD) Server Token: e7 9a 6d 51  [Saved to be used later on]
''''//(DWORD) UDP Token: 09 29 37 00 [Saved to be used later on in 0x09]
''''//(QWORD) Filetime: 00 8b 51 03 70 5f c7 01  [timestamp for version check used in BNLS 0x1A]
''''//(String) filename: 6c 6f 63 6b 64 6f 77 6e 2d 49 58 38 36 2d 31 38 2e 6d 70 71 00[mpq filename, terminated by a null byte]
''''//(String) b4 ce 3d dc 82  53 16 61 3e 05 23 5a 3b bb e5 92 00  [formula string used to do check revision, terminated by a null byte]

Chat.AddChat vbGreen, "Recieved 0x50"


             If buf2.GetDWORD(Mid(Data, 5, 4)) = &H0 Then 'passed our client's logon challenge
                 
                'now to parse out the important infomation from the packet
                 Dim tmpFileTime As String * 8

                'this is used later on, so we must remember it, thats why the "ServerToken" String is declared publicly in the modPkts
                
                modPkts.ServerToken = buf2.GetDWORD(Mid(Data, 9, 4)) ' e7 9a 6d 51
                
                tmpFileName = buf2.GetSTRING(Mid(Data, 25))  'will look for a null byte as from 25th byte on and cut out the string when it finds the null byte
                
                tmpFileTime = Mid$(Data, 17, 8) 'used in BNLS 0x1A

                tmpFormula = buf2.GetSTRING(Mid(Data, 26 + Len(tmpFileName))) 'get the string behind the filename string (using the length of the filename)
                
                modPkts.UDPToken = buf2.GetDWORD(Mid(Data, 13, 4)) 'To be sent 3x via 0x09UDP
                'Call modPkts.Send0x09 '///call 3x 0x09 UDP
       
       
                    Dim Checksum As Long
                    Dim exeversion As Long
                    Dim vidBuf() As Byte
                    Dim EXEInfo As String * 128 'Allocate enough space
                    Dim str As String
                    
                    
                    tmpFileName = Replace(tmpFileName, ".mpq", ".dll") 'Need to replace the .mpq extension with .dll for libbnet
    
                    vidBuf = LoadResData(101, "CUSTOM") 'Loads Star.bin data from the resource file
 
                    str = StrConv(vidBuf, vbUnicode)  'Need to convert the byte array into Unicode format to be read properly
 
 '///Make a function here to check if user even has Hash files in folder or Lockdown folder

If InStr(1, tmpFileName, "lockdown") = 1 Then
    If checkrevision_ld_raw_video(App.Path & "\STAR\starcraft.exe", App.Path & "\STAR\storm.dll", App.Path & "\STAR\battle.snp", tmpFormula, exeversion, Checksum, EXEInfo, App.Path & "\Lockdown\" & tmpFileName, str, Len(str)) = 0 Then
     Chat.AddChat vbRed, "CheckRevision Failed"
     sckBnet.Close
      Exit Sub
    Else
       Chat.AddChat vbGreen, "CheckRevision Passed!"
    End If
Else
    If checkrevision(App.Path & "\STAR\starcraft.exe", App.Path & "\STAR\storm.dll", App.Path & "\STAR\battle.snp", tmpFormula, exeversion, Checksum, EXEInfo, tmpFileName) = 0 Then
     Chat.AddChat vbRed, "CheckRevision Failed"
     sckBnet.Close
      Exit Sub
    Else
       Chat.AddChat vbGreen, "CheckRevision Passed!"
    End If
End If

                Call modPkts.Send0x51(tmpFormula, tmpFileName, Checksum, exeversion, EXEInfo)     ''We now send the next packet in the sequence 0x51, and also send it the two stored variables we got from 0x50
                        
            Else 'failed challenge (unless this is w3..)
            
                        Chat.AddChat vbRed, "Failed 0x50 Logon"
                        sckBnet.Close 'Close Connection
            
            
            End If
        
    
Case &H51

Chat.AddChat vbGreen, "Recieved 0x51"

    Select Case buf2.GetDWORD(Mid(Data, 5, 4))
        Case &H0
        Chat.AddChat vbGreen, "CDKey/Version Good"
            Call modPkts.Send0x3A
        Case &H100
            Chat.AddChat vbRed, "Game version recognized, but out of date!"
            sckBnet.Close 'Close Connection
        Case &H101
            Chat.AddChat vbRed, "Game version unrecognized!"
            sckBnet.Close 'Close Connection
        Case &H102
            Chat.AddChat vbRed, "Game version must be downgraded!"
            sckBnet.Close 'Close Connection
        Case &H200
            Chat.AddChat vbRed, "Invalid CD-key!"
            sckBnet.Close 'Close Connection
        Case &H203
            Chat.AddChat vbRed, "Wrong product!"
            sckBnet.Close 'Close Connection
        Case &H202
            Chat.AddChat vbRed, "CD-key banned by Battle.net!"
            sckBnet.Close 'Close Connection
        Case &H201
            Chat.AddChat vbRed, "CD-key in use!"
            sckBnet.Close 'Close Connection
    End Select
    

Case &H3A

Chat.AddChat vbGreen, "Recieved 0x3A"

    Select Case buf2.GetDWORD(Mid(Data, 5, 4))
        Case &H0
        Dim udpstamp As Long
            Chat.AddChat vbGreen, "Logon passed!"
            Call modPkts.Send0x14
            Call modPkts.Send0x0A
            'Call modPkts.Send0x0B
            Call modPkts.Send0x0C
        Case &H1
            Chat.AddChat vbRed, "Account does not exist!"
            Call modPkts.Send0x3D
        Case &H2
            Chat.AddChat vbRed, "Invalid password!"
            sckBnet.Close 'Close Connection
        Case &H6
            Chat.AddChat vbRed, "Account closed!"
            sckBnet.Close 'Close Connection
        Case Else
            Chat.AddChat vbRed, "Failed 0x3A Logon"
            sckBnet.Close 'Close Connection
    End Select


Case &H3D 'Creating Account
Chat.AddChat vbGreen, "Recieved 0x3D"
If buf2.GetDWORD(Mid(Data, 5, 4)) = &H0 Then  'Account Created
Chat.AddChat vbGreen, IniName & " was created!"
Call modPkts.Send0x3A
Else

If buf2.GetDWORD(Mid(Data, 5, 4)) = &H2 Then 'invalid
Chat.AddChat vbRed, IniName & " has invalid characters!"
Exit Sub
Else


If buf2.GetDWORD(Mid(Data, 5, 4)) = &H3 Then 'Name contained bad stuff
Chat.AddChat vbRed, IniName & " contains bad words!"
Exit Sub
Else

If buf2.GetDWORD(Mid(Data, 5, 4)) = &H4 Then 'Account Exists
Chat.AddChat vbRed, IniName & " already exists!"
Exit Sub

Else


If buf2.GetDWORD(Mid(Data, 5, 4)) = &H6 Then 'Account doesnt have enough alphanumeric
Chat.AddChat vbRed, IniName & " doesn't have enough characters!"
Exit Sub

End If
End If
End If
End If
End If


Case &H25
    
    Call modPkts.Send0x25(Data)
    
    Case &H9 ' Game List
   ' Chat.AddChat vbWhite, "Recieved 0x09"
    Dim GameListData() As String
    Dim ParseData As String
    Dim SpltData  As String
    Dim NumberofGames As Long
    Dim C As Integer
    Dim strTempData As String
    Dim GameType As String
    Dim Port As Long
    Dim IP As String
    Dim TimeofGame As Long
    Dim Gamename As String
    Dim hostname As String
    Dim MapName As String
Dim GameTime As Long
    ParseData = Mid(Data, 9, Len(Data))
    NumberofGames = buf2.MakeLong(Mid(Data, 5, 4))
     'Chat.AddChat vbRed, "Amount of Games Listed: " & NumberofGames
    
    
      If buf2.GetDWORD(Mid(Data, 5, 4)) = &H0 Then   'UhOh Number of games is 0
      
      Dim WatID As Byte
          WatID = Asc(Mid(Data, 9, 1))

      Select Case WatID
      Case &H0 'Game is here and good
      Chat.AddChat vbGreen, "Game is here"
      
      Case &H1 'Game doesnt exist
      'Chat.AddChat vbRed, "Game doesn't exist"
      Call modPkts.Send0x09TCP
      KillTimer frmMain.hWnd, 1000
      Case &H2 'Incorrect Pw
      'Chat.AddChat Index, "Incorrect PW for Game", &HFF00FF
      
      Case &H3 'Game is full
      'Chat.AddChat Index, "Game is full", &HFF00FF
      
      Case &H4 'Game already started
      'Chat.AddChat Index, "Game has already started", &HFF00FF
      
      Case &H6 'Too many server requests
      Chat.AddChat vbRed, "Too Many Server Requests"
      
      End Select

Else
     ' Chat.AddChat vbWhite, "Retrieving GameList"
Call SetTimer(frmMain.hWnd, 1000, 10000, AddressOf TimerProc)

Dim TypeID As Byte
TypeID = Mid$(buf2.GetWORD(ParseData), 1, 2)
 
 GameListData = Split(ParseData, Chr$(&HD) & Chr(&H0))  'Hex(TypeID)
 
Debug.Print "Ubound = " & UBound(GameListData)

Dim LoopMe As Long

For LoopMe = 0 To UBound(GameListData)

 Gamename = buf2.GetSTRING(Mid$(GameListData(LoopMe), 33))
 GameTime = buf2.MakeLong(Mid(GameListData(LoopMe), 29, 4))
 MapName = buf2.GetSTRING(Mid$(StrReverse(buf2.GetSTRING0D(Mid$(StrReverse(GameListData(LoopMe)), 1))), 1))
 hostname = StrReverse(buf2.GetSTRING2C(Mid$(StrReverse(GameListData(LoopMe)), Len(MapName) + 2)))
MapName = modFunctions.ClearedMapData(MapName)
Debug.Print Time & " -------------------"
Debug.Print "[Game = " & Gamename & "] [Host = " & hostname & "] [Map= " & MapName & "]"
Debug.Print "Gametime = " & GameTime
If modFunctions.CheckFilterMatch(Gamename, MapName, hostname) = True And GameTime <= TimeFilter Then
Send0x0E "/me - " & Gamename & " was made by " & hostname & " on " & modFunctions.ClearedMapData(MapName) & " (" & Time & " " & GameTime & " secs ago) -"
Clipboard.SetText Gamename
FillLast5Games Gamename, GameTime
'NotifyAll GameName & " was made by " & hostname & " on " & CommandHandler.ClearedMapData(Mapname) & " (" & Time & " " & GameTime & " secs ago) -"
End If

Next LoopMe

 
End If

    
    
    
    
Case &HF
        ParseBinary Data
End Select




End Sub
Private Sub sckUDP_DataArrival(ByVal bytesTotal As Long)

  Dim Data As String
   Call sckUDP.GetData(Data, vbString)
 
If buf2.GetDWORD(Mid$(Data, 1, 4)) = &H5 Then
Chat.AddChat vbBlue, "Recieved UDP Reponce back from bnet"
udpcount = udpcount + 1
If udpcount > 2 Then
Chat.AddChat vbRed, "UDP 0x05 Count > 4"
udpst = buf2.GetDWORD(Mid$(Data, 5, 4))

'MsgBox udpst

End If
End If
 
End Sub

Private Sub Timer1_Timer()
timrcounter = timrcounter + 1
End Sub

Private Sub tmrflood_timer()
    Dim Message As String
    Message = Queue.Peek
    If LenB(Message) > 0 Then
        If sckBnet.State <> 0 Then
            Send0x0E Queue.Dequeue
        End If
    End If
End Sub
Private Sub txtChan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
ReleaseCapture
SendMessage hWnd, WM_NCLBUTTONDOWN, 2, 0&
End If

End Sub

Private Sub txtChan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape2.BorderWidth = 1
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    'On Error Resume Next
    If KeyAscii = 13 Then
        If Left$(txtSend.Text, 1) = "/" Then
            Queue.Enqueue txtSend.Text
            'CommandHandler.ParseInternalCmd txtSend.Text
        Else
            Send0x0E (txtSend.Text)
        End If
        txtSend.Text = ""
        KeyAscii = 0

    End If
End Sub

Public Sub HandleUnload()
        
End Sub

Private Sub txtSend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape2.BorderWidth = 2

End Sub

Private Sub Trayicon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'###########################################################
' This Function Tells what Event happened to icon in
' system tray
'###########################################################

    Dim MSG As Long
    MSG = (X And &HFF) * &H100

    Select Case MSG
        Case 0 'mouse moves
        
        Case &HF00  'left mouse button down
        
        Case &H1E00 'left mouse button up
        
        Case &H3C00  'right mouse button down
        'PopupMenu mnuOP, 2, , , mnuCTRL 'show the popoup menu
        
        Case &H2D00 'left mouse button double click
        NoSysIcon True    'Show App on double clicking Mouse's Left Button
        
        Case &H4B00 'right mouse button up
        
        Case &H5A00 'right mouse button double click
        
    End Select
   
End Sub
 
'###########################################################
' This Function Show TrayIcon.Picture in System Tray
' Don't Bother what it says
' To change TrayIcon's ToolTip goto 2nd Last Line
'###########################################################
Public Function ShowProgramInTray()
INTRAY = True   'Means App is now in Tray
    
    NI.cbSize = Len(NI) 'set the length of this structure
    NI.hWnd = TrayIcon.hWnd 'control to receive messages from
    NI.uID = 0 'uniqueID
    NI.uID = NI.uID + 1
    NI.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP 'operation flags
    NI.uCallbackMessage = WM_MOUSEMOVE 'recieve messages from mouse activities
    NI.hIcon = TrayIcon.Picture  'the location of the icon to display
  
' Change System Tray Icon's Tool Tip Here bt don't delete chr$(0) [its line carriage here]
    
    NI.szTip = "Game Announcer - Created by Myst and cHip" + Chr$(0) 'LoadResString(Language) + Chr$(0)  'the tool tip to display"
    result = Shell_NotifyIconA(NIM_ADD, NI)     'add the icon to the system tray
End Function


'###########################################################
' This Function Delete TrayIcon.Picture from System Tray
' Don't Bother what it says
'###########################################################
Private Sub DeleteIcon(pic As Control)
INTRAY = False  'Means app is unloaded or Max mode
    
    ' On remove, we only have to give enough information for Windows
    ' to locate the icon, then tell the system to delete it.
    NI.uID = 0 'uniqueID
    NI.uID = NI.uID + 1
    NI.cbSize = Len(NI)
    NI.hWnd = pic.hWnd
    NI.uCallbackMessage = WM_MOUSEMOVE
    result = Shell_NotifyIconA(NIM_DELETE, NI)
End Sub

'###########################################################
' This Function controls 3 funtions
' 1] Visibility of App Form
' 2] Visibility of Tray Icon
' 3] Menu Caption Control
'###########################################################

Public Function NoSysIcon(maxIcon As Boolean)
    Select Case maxIcon
    Case False   'Case App in Min Mode
        Me.Visible = False
        ShowProgramInTray               'Now show TrayIcon PictureBox's Picture in SysTray as icon
        'mnuCTRL.Caption = "E&xpand Application"
    
    Case Else   'Case App in Max Mode
        Me.Visible = True
        DeleteIcon TrayIcon
        'mnuCTRL.Caption = "Minimize App to System Tray"
    
    End Select

End Function

Public Sub FillLast5Games(Game As String, Time As Long)
    If CloneDetectLast5(Game) <> True Then
If CntThis <= 4 And FirstRollover = True Then
Last5Gamess(CntThis).GameNames = Game
Last5Gamess(CntThis).GameTimes = Time

CntThis = CntThis + 1

Else
FirstRollover = False
Last5Gamess(0).GameNames = Last5Gamess(1).GameNames
Last5Gamess(1).GameNames = Last5Gamess(2).GameNames
Last5Gamess(2).GameNames = Last5Gamess(3).GameNames
Last5Gamess(3).GameNames = Last5Gamess(4).GameNames
Last5Gamess(4).GameNames = Game

Last5Gamess(0).GameTimes = Last5Gamess(1).GameTimes
Last5Gamess(1).GameTimes = Last5Gamess(2).GameTimes
Last5Gamess(2).GameTimes = Last5Gamess(3).GameTimes
Last5Gamess(3).GameTimes = Last5Gamess(4).GameTimes
Last5Gamess(4).GameTimes = Time


CntThis = 0

 
End If
        End If
End Sub

 
Public Function CloneDetectLast5(Game As String) As Boolean
Dim X As Integer

For X = 0 To UBound(Last5Gamess)
    If Game = Last5Gamess(X).GameNames Then
        CloneDetectLast5 = True
    End If
Next X



End Function
Public Function GrabWebTitle(site As String) As String
Dim HoldThis() As String
Dim UTube() As String
Dim hhold As String
Dim Title As String
Dim title2() As String
Dim X As Long
On Error Resume Next
Dim op As Integer
op = 1
 


 



hhold = Inet1.OpenURL(site)


HoldThis() = Split(hhold, "<")

'MsgBox UBound(HoldThis)
Do Until X = UBound(HoldThis)
'MsgBox HoldThis(x)
If InStr(1, HoldThis(X), "title", vbTextCompare) Then
    
     
  ' MsgBox Inet2.RemoteHost
    Title = Replace(HoldThis(X), "title>", "")
    
    
    If Split(Title, "") = True Then
    'UTube() = Split(title2, "")
 If InStr(1, Title, "YouTube") Then
 Dim indent As String
 indent = Len(Mid$(Title, 1, InStr(1, Title, "YouTube")))
Title = Mid$(Title, indent)
Dim utubetitle As String
utubetitle = Mid$(Title, InStr(Title, "-"))


Title = "YouTube " & Trim(Mid$(Title, InStr(Title, "-")))
 'MsgBox buf2.StrToHex(title)
 'MsgBox buf2.HexToStr("FF")
 Title = Replace(Title, buf2.HexToStr("0A"), "")
 Dim xz As Integer
  xz = 1
  
 Do Until xz = Len(Title) + 1
 If Asc(Mid$(Title, xz, 1)) <= "20" Then
 Replace Title, Asc(Mid$(Title, xz, 1)), ""
 End If
 
 xz = xz + 1
 
 
 
 DoEvents
 Loop
 
  End If
     End If
    
    
    
    GrabWebTitle = Title
    If Title = "" Then
    Title = "Unknown title, website may not exist"
    Exit Function
    End If
    'MsgBox title
    
    
   

    Exit Function
End If
    X = X + 1
Loop
DoEvents

End Function

