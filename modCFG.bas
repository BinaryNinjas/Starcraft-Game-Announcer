Attribute VB_Name = "modCFG"
Public Function ReadINI(ByVal appname As String, Key As String) As String
    Dim sFile As String
    Dim sDefault As String
    Dim lSize As Integer
    Dim L As Long
    Dim sUser As String
    sUser = Space$(128)
    lSize = Len(sUser)
    sFile = App.Path & "\Config.ini"
    sDefault = ""
    L = GetPrivateProfileString(appname, Key, sDefault, sUser, lSize, sFile)
    sUser = Mid(sUser, 1, InStr(sUser, Chr(0)) - 1)
    ReadINI = sUser
End Function

Public Sub WriteINI(wiSection$, wiKey$, ByVal wiValue$)
    Dim sFile As String
    sFile = (App.Path & "\Config.ini")
    WritePrivateProfileString wiSection, wiKey, wiValue, sFile
End Sub

Public Sub LoadCFG()
    '''''''''''' reg acct vars ''''''''''''''
    IniName = ReadINI("Connection", "Username")
    IniPass = ReadINI("Connection", "Password")
    IniHome = ReadINI("Connection", "HomeChan")
    IniVerByte = ReadINI("Connection", "VerByte")
    IniCDKey = ReadINI("Connection", "CDKey")
    IniServer = ReadINI("Connection", "Server")
    IniTrigger = ReadINI("Options", "Trigger")
    IniIdleMsg = ReadINI("Options", "IdleMsg")
    strMast = ReadINI("Options", "Master")
    frmMain.sckChat.RemoteHost = IniServer
    frmMain.sckChat.LocalPort = 0
    WatchType = UCase(Trim(ReadINI("Options", "Detect")))


End Sub


