VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSCUDP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''clsSCUDP - Starcraft Game Protocal
''''By Myst
''''www.DarkBlizz.org
''''irc.darkblizz.org #darkblizz
'''''''''''''''''''''''''''''''''''''''

'Class offset at 13
'Class 0 id offset = 14
'Class 1 id offset = 17
'Class 2 id offset = 17


Public Function Send0x01UDP()
'Class 0 - 0x01 - Request to join game (3x)
'Player > Host
'0020  ad 43 17 e0 17 e0 00 1c  18 bb 00 00 00 00 28 c4   .C...... ......(.
'0030  10 00 00 00 01 00 00 01  ff 00 01 00 00 00         ........ ......




End Function

