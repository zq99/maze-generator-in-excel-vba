VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csNextCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

Private mRange          As Range
Private mintDirection   As Integer

Property Get Cell()
    Set Cell = mRange
End Property

Property Get Direction()
    Direction = mintDirection
End Property

Public Sub SetRangeNext(ByVal rng As Range, ByVal iDirection As Integer)
    Set mRange = rng
    mintDirection = iDirection
End Sub




