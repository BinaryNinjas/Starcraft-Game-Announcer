Attribute VB_Name = "Chat"
Public Function AddChat(Color As String, MSG As String, Optional Fade As Boolean)
With frmMain.txtChan
If Len(.Text) > 100000 Then .Text = ""
.SelStart = Len(.Text)
.SelColor = "&H" & StrReverse("33FFCC")
.SelText = "[" & Time & "]"
.SelColor = Color
.SelText = " " & MSG & vbCrLf
.SelStart = Len(.Text)
End With
End Function
Public Function AddChat2(Color As String, MSG As String, Optional Fade As Boolean)
With frmMain.txtChan
If Len(.Text) > 100000 Then .Text = ""
.SelStart = Len(.Text)

.SelColor = Color
.SelText = " " & MSG & vbCrLf
.SelStart = Len(.Text)
End With
End Function
Public Function AddChat3(Color As String, MSG As String, User As String, Optional Fade As Boolean)
With frmMain.txtChan
If Len(.Text) > 100000 Then .Text = ""
.SelStart = Len(.Text)
.SelColor = "&H" & StrReverse("33FFCC")
.SelText = "[" & Time & "]"
.SelColor = vbYellow
.SelText = " <" & User & ">"
.SelColor = Color
.SelText = " " & MSG & vbCrLf
.SelStart = Len(.Text)
End With
End Function
