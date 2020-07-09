VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVisited"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

Private mdicVisited As Scripting.Dictionary

Private Sub Class_Initialize()
    Set mdicVisited = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    Set mdicVisited = Nothing
End Sub

Public Function Count() As Integer
    Count = mdicVisited.Count
End Function

Public Function Add(ByVal rngCell As Range)
    If mdicVisited Is Nothing Then
        Set mdicVisited = New Scripting.Dictionary
    End If
    mdicVisited.Add rngCell.Address, rngCell
End Function

Public Function IsRangeVisited(ByVal rngCell As Range) As Boolean
    IsRangeVisited = mdicVisited.Exists(rngCell.Address)
End Function
