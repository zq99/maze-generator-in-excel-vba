VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

Dim moStack As Object

Private Sub Class_Initialize()
    Set moStack = CreateObject("System.Collections.Stack")
End Sub

Public Function Push(ByVal Value As Range)
    moStack.Push Value
End Function

Public Function Pop() As Variant
    Pop = moStack.Pop
End Function

Public Function Top() As Range
    Dim arr As Variant
    arr = moStack.ToArray()
    Set Top = arr(LBound(arr))
End Function

Private Sub Class_Terminate()
    If (Not moStack Is Nothing) Then
        moStack.Clear
        Set moStack = Nothing
    End If
End Sub
