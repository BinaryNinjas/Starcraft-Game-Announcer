Attribute VB_Name = "modVars"
Option Explicit

'' Basic account vars ''
Public IniName As String, IniPass As String, IniHome As String, IniDelay As String, IniServer As String, IniVerByte As String, IniTrigger As String, IniCDKey As String
Public Hasops As Boolean, UserDesignated As String, IniShitMsg As String, IniIdleMsg As String, lockDown As Boolean

''  Global Classes ''
 Public CommandHandler As clsCommandHandler
Public Queue As clsQueue
Public UserList As Collection
Public strMast As String
Public trigger As String

''  Color Variables ''
Public clrTimeStamp As Long, clrBotInfo As Long, clrTxtInfo As Long
Public clrUserInfo As Long, clrLVOps As Long, clrLVUser As Long
Public clrAlert As Long, clrUserHeadings As Long, clrChatHeadings As Long
Public clrOpHeadings As Long, clrChat As Long, clrEmote As Long

''  Toggle Variables ''
Public IniMaximized As Boolean


''  Status Variables ''
Public curChan As String

Public Const ID_USER = &H1
Public Const ID_JOIN = &H2
Public Const ID_LEAVE = &H3
Public Const ID_WHISPFROM = &H4
Public Const ID_TALK = &H5
Public Const ID_BROADCAST = &H6
Public Const ID_CHAN = &H7
Public Const ID_FLAGS = &H9
Public Const ID_WHISPTO = &HA
Public Const ID_CHANFULL = &HD
Public Const ID_CHANDONTEXIST = &HE
Public Const ID_CHANRESTRICTED = &HF
Public Const ID_INFO = &H12
Public Const ID_ERROR = &H13
Public Const ID_EMOTE = &H17
Public Const SID_FRIENDSLIST = &H65
Public FirstTime As Boolean
