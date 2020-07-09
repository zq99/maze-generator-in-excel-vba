Attribute VB_Name = "mdMaze"
'*********************************************************
' Maze Generation using a recursive backtracker algorithm
'*********************************************************

Option Explicit

Private Const CINT_COLOUR_MAZE          As Integer = 24
Private Const CINT_COLOUR_BLACK         As Integer = 1
Private Const CINT_COLOUR_CURRENT_CELL  As Integer = 3

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)            'For 32 Bit Systems
#End If


Public Sub Start()

    Dim moVisited              As New clsVisited
    Dim moStack                As New clsStack
    Dim oNextCell              As csNextCell
    Dim coNextMoveOptions      As Collection
    Dim rngCurrent             As Range
    Dim rngCanvas              As Range
    Dim varDirection           As Variant
    Dim intIndex               As Integer
    Dim lRow                   As Long
    Dim iCol                   As Long
    
    Call ResetMaze
    varDirection = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
    Set rngCanvas = Range("canvas")
    With rngCanvas
      lRow = GetRandomNumber(.Rows.Count + .Row - 1, .Row)
      iCol = GetRandomNumber(.Columns.Count + .Column - 1, .Column)
    End With
    Set rngCurrent = rngCanvas.Parent.Cells(lRow, iCol)
    
    Do While Not IsMazeCompleted(rngCanvas, moVisited)
        DoEvents
        Set coNextMoveOptions = New Collection
        For intIndex = LBound(varDirection) To UBound(varDirection)
            Set oNextCell = GetNextCellCandidate(rngCurrent, varDirection(intIndex), moVisited, rngCanvas)
            If Not oNextCell Is Nothing Then
                coNextMoveOptions.Add oNextCell
            End If
        Next
        
        If coNextMoveOptions.Count = 0 Then
            moStack.Pop
            Call HighlightCell(False, rngCurrent)
            Set rngCurrent = moStack.Top
            Call HighlightCell(True, rngCurrent)
        Else
            Set oNextCell = coNextMoveOptions.Item(GetRandomNumber(coNextMoveOptions.Count, 1))
            Call RemoveBorder(rngCurrent, oNextCell.Direction)
            Call HighlightCell(False, rngCurrent)
            Set rngCurrent = oNextCell.Cell
            Call HighlightCell(True, rngCurrent)
            moVisited.Add rngCurrent
            moStack.Push rngCurrent
        End If
    Loop
    
    Set rngCurrent = Nothing
    Set moStack = Nothing
    Set moVisited = Nothing
    Set oNextCell = Nothing
    Set coNextMoveOptions = Nothing
    Set rngCanvas = Nothing

End Sub

Public Sub ResetMaze()
    Call ResetLines(shtMaze.Range("canvas"))
End Sub

Private Function GetNextCellCandidate(ByVal rngCurrent As Range, _
                                      ByVal intDirection As Integer, _
                                      ByVal oVisited As clsVisited, _
                                      ByVal rngArea As Range) As csNextCell
    Dim lngRow    As Long
    Dim intCol    As Integer
    Dim oNext     As csNextCell
    
    lngRow = rngCurrent.Row
    intCol = rngCurrent.Column

    Select Case intDirection
    Case xlEdgeLeft
        intCol = intCol - 1
    Case xlEdgeRight
        intCol = intCol + 1
    Case xlEdgeTop
        lngRow = lngRow - 1
    Case xlEdgeBottom
        lngRow = lngRow + 1
    End Select
    
    If IsValidExcelCell(lngRow, intCol) Then
        If IsInBoardRange(rngCurrent.Parent.Cells(lngRow, intCol), rngArea) Then
            If Not oVisited.IsRangeVisited(rngCurrent.Parent.Cells(lngRow, intCol)) Then
                Set oNext = New csNextCell
                oNext.SetRangeNext rngCurrent.Parent.Cells(lngRow, intCol), intDirection
                Set GetNextCellCandidate = oNext
                Set oNext = Nothing
                Exit Function
            End If
        End If
    End If
    
    
End Function

Private Function IsInBoardRange(ByVal rngCell As Range, ByVal rngBoard) As Boolean
    If Application.Intersect(rngCell, rngBoard) Is Nothing Then
        IsInBoardRange = False
    Else
        IsInBoardRange = True
    End If
End Function

Private Function IsValidExcelCell(ByVal lngRow As Long, ByVal intCol As Integer) As Boolean
    IsValidExcelCell = IIf(lngRow >= 1 And intCol >= 1, True, False)
End Function

Private Function IsMazeCompleted(ByVal rngBoard As Range, ByVal oVisited As clsVisited)
    IsMazeCompleted = IIf(rngBoard.Cells.Count = oVisited.Count, True, False)
End Function

Private Function GetRandomNumber(ByVal intUpper As Integer, ByVal intLower As Integer) As Integer
    Randomize
    GetRandomNumber = Int((intUpper - intLower + 1) * Rnd + intLower)
End Function

Private Sub ResetLines(ByVal rngArea As Range)
    Dim varLines    As Variant
    Dim intLines    As Integer
    
    Application.ScreenUpdating = False
    varLines = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
    For intLines = LBound(varLines) To UBound(varLines)
        With rngArea.Borders(varLines(intLines))
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        rngArea.Interior.Pattern = xlNone
        rngArea.Interior.ColorIndex = CINT_COLOUR_BLACK
    Next
    Application.ScreenUpdating = True
    
End Sub

Private Sub RemoveBorder(ByVal rngArea As Range, ByVal intDirection)
    rngArea.Borders(intDirection).LineStyle = xlNone
End Sub

Private Sub HighlightCell(ByVal fill As Boolean, ByVal rng As Range)
    Sleep 1
    With rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        If fill Then
            .ColorIndex = CINT_COLOUR_CURRENT_CELL
        Else
            .ColorIndex = CINT_COLOUR_MAZE
        End If
    End With
End Sub

