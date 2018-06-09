VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub BeginParse(strdata)
    Dim strStrings() As String, I
    SplitStr strdata, strStrings, vbCrLf
    For I = 1 To UBound(strStrings)
        Parse strStrings(I)
    Next I
End Sub
Public Sub Parse(ByVal strdata As String)
    On Error Resume Next
    Dim strData1, strData2, I
    Dim strArgs() As String
    If Left(strdata, 1) = Chr(10) Then
        strdata = Right(strdata, Len(strdata) - 1)
    End If
    If Val(Left(strdata, 4)) < 1001 Or Val(Left(strdata, 4)) > 3000 Then
        Exit Sub
    End If
    I = InStr(1, strdata, Chr(34), vbBinaryCompare)
    If I <> 0 Then
        strData1 = Left(strdata, I - 2)
        strData2 = Mid(strdata, I + 1, Len(strdata) - I - 1)
    Else
        strData1 = strdata
        strData2 = ""
    End If
    
    SplitStr strData1, strArgs(), " "
    
    Select Case Val(strArgs(1))
        Case 1005
            EID_Talk strArgs(3), strArgs(4), strData2
            Exit Sub
        Case 1023
            EID_Emote strArgs(3), strArgs(4), strData2
            Exit Sub
        Case 1007
            EID_Channel strData2
            Exit Sub
        Case 1001
            EID_User strArgs(3), strArgs(4), strArgs(5)
            Exit Sub
        Case 1002
            EID_Join strArgs(3), strArgs(4), strArgs(5)
            Exit Sub
        Case 1003
            EID_Leave strArgs(3), strArgs(4)
            Exit Sub
        Case 1010
            EID_SendWhisper strArgs(3), strArgs(4), strData2
            Exit Sub
        Case 1004
            EID_RecvWhisper strArgs(3), strArgs(4), strData2
            Exit Sub
        Case 1016, 1018, 300
            EID_Info strData2
            Exit Sub
        Case 1009
            EID_Flags strArgs(3), strArgs(4)
            Exit Sub
        Case 1006
            EID_Broadcast strData2
            Exit Sub
        Case 1019
            EID_Error strData2
            Exit Sub
        Case 2010
            EID_Name strArgs(3)
            Exit Sub
        Case 1013
            EID_ChannelFull strData2
            Exit Sub
        Case 1014
            EID_ChannelNotExist strData2
            Exit Sub
        Case 1015
            EID_ChannelRestricted strData2
            Exit Sub
        Case 2000
            EID_Null strData2
            Exit Sub
        Case Else
            EID_Unknown strdata
            Exit Sub
    End Select
End Sub
Private Sub EID_Unknown(strText)
    AddTimeStamp frmMain.txtChan
    If Left$(strText, 15) = "Login incorrect" Then
        AddRTB frmMain.txtChan, clrBotInfo, "Your Login is Incorrect"
    Else
        AddRTB frmMain.txtChan, clrBotInfo, strText
    End If
End Sub

Private Sub EID_User(strUser, strFlag, strProduct)
    Dim s As Long
    s = CLng("&H" & strFlag)
    If (PrepareCheck(strUser) = PrepareCheck(IniName)) And ((s = 2) Or (s = 18)) Then
        hasOps = True
    End If
End Sub

Private Sub EID_Join(strUser, strFlag, strProduct)
    frmMain.lvChan.ListItems.Add frmMain.lvChan.ListItems.Count + 1, , strUser, , GetIconCode(strProduct, CLng("&H" & strFlag))
    frmMain.txtCLabel.text = curChan & " (" & frmMain.lvChan.ListItems.Count & ")"
    If Mid$(strFlag, 4, 1) = "1" Or Mid$(strFlag, 4, 1) = "2" Then
        frmMain.lvChan.ListItems.Item(frmMain.lvChan.ListItems.Count).ForeColor = clrLVOps
    Else
        frmMain.lvChan.ListItems.Item(frmMain.lvChan.ListItems.Count).ForeColor = clrLVUser
    End If
    If IniShowEIDJoinLeave Then
        AddTimeStamp frmMain.txtChan
        AddRTB frmMain.txtChan, clrUserInfo, strUser & " has joined the channel using " & strProduct & "."
    End If
    If hasOps = True And lockDown = True Then
        If Database.CheckSafelist(PrepareCheck(strUser)) = False Then
            Queue.Enqueue "/ban " & strUser & " Lockdown is on." & vbCrLf
        End If
    End If
End Sub


Private Sub EID_Leave(strUser, strFlag)
    frmMain.lvChan.ListItems.Remove (frmMain.lvChan.FindItem(strUser).index)
    frmMain.txtCLabel.text = curChan & " (" & frmMain.lvChan.ListItems.Count & ")"
    If IniShowEIDJoinLeave Then
        AddTimeStamp frmMain.txtChan
        AddRTB frmMain.txtChan, clrUserInfo, strUser & " has left the channel."
    End If
End Sub

Private Sub EID_RecvWhisper(strUser, strFlag, strText)
   ' Dim i As Integer
   ' For i = 1 To colWBox.Count
   '     If colWBox.Item(i).Visible = True Then
   '     If PrepareCheck(strUser) = PrepareCheck(colWBox.Item(i).Caption) Then
    '        colWBox.Item(i).HandleWBox "[" & Time & "] " & "<" & strUser & "> ", strText & vbCrLf, 0
   '         Exit Sub
   '     End If
   '     End If
   ' Next i
    AddTimeStamp frmMain.txtWhisper
    AddRTB frmMain.txtWhisper, clrChatHeadings, "<From: " & strUser & "> ", clrWhisperTxt, strText
    If Left$(strText, 1) = IniTrigger Then
        CommandHandler.ParseCommand strUser, Mid$(strText, 2, Len(strText) - 1)
    End If
End Sub

Private Sub EID_Talk(strUser, strFlag, strText)
   ' If IniFilter = 1 Then
    '    Dim i As Integer
    '    For i = 0 To UBound(FilterList)
    '        If Battlenet2.PrepareCheck(strText) Like Battlenet2.PrepareCheck(FilterList(i)) Then
   '             Exit Sub
    ''        End If
    '    Next i
   ' End If
    AddTimeStamp frmMain.txtChan
    If Mid$(strFlag, 4, 1) = "2" Or Mid$(strFlag, 4, 1) = "1" Then
        AddRTB frmMain.txtChan, clrOpHeadings, "<" & strUser & "> ", clrChat, strText
    Else
        AddRTB frmMain.txtChan, clrChatHeadings, "<" & strUser & "> ", clrChat, strText
    End If
    
    If Left$(strText, 1) = IniTrigger Then
        CommandHandler.ParseCommand strUser, Mid$(strText, 2, Len(strText) - 1)
    End If
    If PrepareCheck(strText) = "?trigger" Or _
        PrepareCheck(strText) = "?trig" Or _
        PrepareCheck(strText) = "?t" Then
        CommandHandler.ParseCommand strUser, Mid$(strText, 2, Len(strText) - 1)
    End If
End Sub

Private Sub EID_Broadcast(strText)
    AddTimeStamp frmMain.txtChan
    AddRTB frmMain.txtChan, clrAlert, strText & vbCrLf
End Sub

Private Sub EID_Channel(strText)
    hasOps = False
    UserDesignated = ""
    AddTimeStamp frmMain.txtChan
    AddRTB frmMain.txtChan, clrUserInfo, "Joined Channel: " & strText
    curChan = strText
    frmMain.Caption = " Game Announcer " & App.Major & "." & App.Minor & " (Build: " & App.Revision & ") User: [" & IniName & "] on server " & IniServer
    'If (IniOpLog = 1) And (opBool = True) And (PrepareCheck(strText) = PrepareCheck(IniHome)) Then
     '   opBool = False
     '   frmBot.tmrOp.Enabled = True
    '    frmBot.HandleQueue "/who " & IniOpChan1
    '    frmBot.HandleQueue "/who " & IniOpChan2
   ' End If

End Sub

Private Sub EID_Flags(strUser, strFlag)
    Dim s As Long
    s = CLng("&H" & strFlag)
    Dim I As Integer
    On Error Resume Next

    
    For I = 0 To frmMain.lvChan.ListItems.Count
        If frmMain.lvChan.ListItems(I).text = strUser Then
            If strFlag = "0010" Then
                frmMain.lvChan.ListItems.Remove I
                frmMain.lvChan.ListItems.Add I, , strUser, , 2
                frmMain.lvChan.ListItems(I).ForeColor = clrLVUser
            Else
                frmMain.lvChan.ListItems.Remove I
                frmMain.lvChan.ListItems.Add I, , strUser, , GetFlagCode(s)
                If GetFlagCode(s) = 2 Then
                    frmMain.lvChan.ListItems(I).ForeColor = clrLVUser
                Else
                    frmMain.lvChan.ListItems(I).ForeColor = clrLVOps
                End If
            End If
        End If
    Next I
    If (PrepareCheck(strUser) = PrepareCheck(IniName)) Then
        If ((s = 2) Or (s = 18)) Then
            hasOps = True
        Else
            hasOps = False
        End If
    End If
End Sub

Private Sub EID_SendWhisper(strUser, strFlag, strText)
   ' Dim i As Integer
   ' For i = 1 To wboxCount
   '     If colWBox.Item(i).Visible = True Then
   '     If PrepareCheck(strUser) = PrepareCheck(colWBox.Item(i).Caption) Then
    '        colWBox.Item(i).HandleWBox "[" & Time & "] " & "<" & frmBot.strName & "> ", strText & vbCrLf, 1
    '        Exit Sub
   '     End If
    '    End If
   ' Next i
    AddTimeStamp frmMain.txtWhisper
    AddRTB frmMain.txtWhisper, clrUserHeadings, "<To: " & strUser & "> ", clrWhisperTxt, strText
End Sub

Private Sub EID_ChannelFull(strText)
    AddTimeStamp frmMain.txtChan
    AddRTB frmMain.txtChan, clrAlert, "Channel " & strText & " is full"
End Sub

Private Sub EID_ChannelNotExist(strText)
    AddTimeStamp frmMain.txtChan
    AddRTB frmMain.txtChan, clrAlert, "Channel " & strText & " does not exist"
End Sub

Private Sub EID_ChannelRestricted(strText)
    AddTimeStamp frmMain.txtChan
    AddRTB frmMain.txtChan, clrAlert, "Channel " & strText & " is restricted"
End Sub

Private Sub EID_Info(strText)
    AddTimeStamp frmMain.txtChan
    AddRTB frmMain.txtChan, clrTxtInfo, strText
   ' If (IniOpLog = 1) And (frmBot.tmrOp.Enabled = True) Then
   '     WriteLogEntry strText
   ' End If
End Sub

Private Sub EID_Error(strText)
    AddTimeStamp frmMain.txtChan
    AddRTB frmMain.txtChan, clrBotInfo, strText
End Sub

Private Sub EID_Emote(strUser, strFlag, strText)
    'If IniFilter = 1 Then
    '    Dim i As Integer
    '    For i = 0 To UBound(FilterList)
     '       If Battlenet2.PrepareCheck(strText) Like Battlenet2.PrepareCheck(FilterList(i)) Then
     '           Exit Sub
     '       End If
    '    Next i
  ' End If
    AddTimeStamp frmMain.txtChan
    AddRTB frmMain.txtChan, clrEmote, "<" & strUser & " " & strText & ">"
      
End Sub

Private Sub EID_Name(strUser)
    frmMain.Caption = " Game Announcer " & App.Major & "." & App.Minor & " (Build: " & App.Revision & ") User: [" & IniName & "] on server " & IniServer
    AddTimeStamp frmMain.txtChan
    AddRTB frmMain.txtChan, clrBotInfo, "Successfully logged in as " & strUser
End Sub
Private Sub EID_Null(strText)
    
End Sub




