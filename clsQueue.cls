VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private colQueue As New Collection

Public Function Dequeue() As String
    If (colQueue.Count > 0) Then
        Dequeue = colQueue.Item(1)
        colQueue.Remove 1
    Else
        Dequeue = vbNullString
    End If
End Function

Public Function Peek() As String
    If (colQueue.Count > 0) Then
        Peek = colQueue.Item(1)
    Else
        Peek = vbNullString
    End If
End Function

Public Function Count() As Long
    Count = colQueue.Count
End Function

Public Sub Clear()
    Set colQueue = New Collection
End Sub
Public Sub Enqueue(ByVal strValue As String)
    If (LenB(strValue) > 0) Then
        colQueue.Add strValue
    End If
End Sub


