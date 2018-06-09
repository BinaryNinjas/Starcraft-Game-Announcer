VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub LoadDatabase()
    Dim strInput As String, strArgs() As String
   
    Open App.Path & "\database.txt" For Input As #1
    
    Do While Not EOF(1)
        Line Input #1, strInput
        strArgs = Split(strInput, ";")
        
        If UBound(strArgs) = 1 Then
        If Not UserExists(strArgs(0)) Then
            Dim X As clsUser
            Set X = New clsUser
            X.Init strArgs(0), strArgs(1)
            UserList.Add X
            Set X = Nothing
        End If
        End If
    Loop
    
    Close #1
    Dim strMast As String
    strMast = ReadINI("Options", "Master")
    mastFlags = getMastFlags()
    If UserExists(strMast) Then
        RemoveDatabase (strMast)
        AddDatabase strMast, mastFlags
    Else
        AddDatabase strMast, mastFlags
    End If
        
End Sub
Public Sub SaveDatabase(User As String, ByVal AccessLevel As String) '
    Dim intFreeFile As Integer, I As Integer
    Const DELIMETER As String = ";"
    
    intFreeFile = FreeFile
    
    Open (App.Path & "\database.txt") For Output As #intFreeFile
    
    For I = 1 To UserList.count
        Print #intFreeFile, UserList.Item(I).User & DELIMETER & UserList.Item(I).Flags
    Next I
    
    Close #intFreeFile
End Sub
Public Function AddDatabase(strUser As String, ByVal strFlags As String) As Boolean
    If UserExists(strUser) = True Then
        AddDatabase = False
    Else
        Dim X As clsUser
        Set X = New clsUser
        X.Init strUser, strFlags
        UserList.Add X
        Set X = Nothing
        AddDatabase = True
    End If
End Function
 
Public Function CheckOwner(strUser As String) As Boolean
    If PrepareCheck(strUser) = PrepareCheck(strMast) Then
        CheckOwner = True
    Else
        CheckOwner = False
    End If
End Function
Public Function CheckMaster(strUser As String) As Boolean
    Dim I As Integer
    CheckMaster = False
    For I = 1 To UserList.count
        If (PrepareCheck(UserList.Item(I).Flags) Like PrepareCheck("*m*")) And (PrepareCheck(strUser) = PrepareCheck(UserList.Item(I).User)) Then
            CheckMaster = True
        End If
    Next I
End Function
Public Function CheckOperator(strUser As String) As Boolean
    Dim I As Integer
    CheckOperator = False
    For I = 1 To UserList.count
        If (PrepareCheck(UserList.Item(I).Flags) Like PrepareCheck("*a*")) And (PrepareCheck(strUser) = PrepareCheck(UserList.Item(I).User)) Then
            CheckOperator = True
        End If
    Next I
End Function
Public Function CheckSafelist(ByVal strUser As String) As Boolean
    Dim I As Integer
    CheckSafelist = False
    For I = 1 To UserList.count
        If (PrepareCheck(strUser) Like PrepareCheck(UserList.Item(I).User)) Or (PrepareCheck(strUser) = PrepareCheck(UserList.Item(I).User)) Then
            If PrepareCheck(UserList.Item(I).Flags) Like PrepareCheck("*s*") Then
                CheckSafelist = True
            End If
        End If
    Next I
End Function
Public Function CheckShitlist(ByVal strUser As String) As Boolean
    Dim I As Integer
    CheckShitlist = False
    
    For I = 1 To UserList.count
        If (PrepareCheck(strUser) Like PrepareCheck(UserList.Item(I).User)) Or (PrepareCheck(strUser) = PrepareCheck(UserList.Item(I).User)) Then
            If PrepareCheck(UserList.Item(I).Flags) Like PrepareCheck("*z*") Then
                CheckShitlist = True
            End If
        End If
    Next I
    If CheckSafelist(strUser) = True Then
        CheckShitlist = False
    End If
    
End Function
Public Function Find(UserTag As String) As String()
    Dim users() As String
    ReDim users(0)
    Dim I As Integer
    For I = 1 To UserList.count
        
                
        If PrepareCheck(UserList.Item(I).User) Like PrepareCheck(UserTag) Or _
           PrepareCheck(UserTag) = PrepareCheck(UserList.Item(I).User) Then
            users(UBound(users)) = UserList.Item(I).User & ";" & UserList.Item(I).Flags
            ReDim Preserve users(UBound(users) + 1)
        End If
    Next I
    Find = users
End Function
Public Function UserExists(strUser As String) As Boolean
    Dim I As Integer
    UserExists = False
    For I = 1 To UserList.count
        If PrepareCheck(UserList.Item(I).User) = PrepareCheck(strUser) Then
            UserExists = True
        End If
    Next I
End Function
Public Function FindFlag(flag As String) As String()
    Dim users() As String
    ReDim users(0)
    Dim I As Integer
    For I = 1 To UserList.count
        If PrepareCheck(UserList.Item(I).Flags) Like PrepareCheck("*" & flag & "*") Then
            users(UBound(users)) = UserList.Item(I).User & ";" & UserList.Item(I).Flags
            ReDim Preserve users(UBound(users) + 1)
        End If
    Next I
    FindFlag = users
End Function

Public Function HighestAccess(strUser As String) As Integer
    If CheckOwner(strUser) Then
        HighestAccess = 3
    ElseIf CheckMaster(strUser) Then
        HighestAccess = 2
    ElseIf CheckOperator(strUser) Then
        HighestAccess = 1
    Else
        HighestAccess = 0
    End If
End Function


