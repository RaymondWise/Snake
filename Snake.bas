Attribute VB_Name = "Snake"
Option Explicit
    Const UP_CODE As Long = 233
    Const DOWN_CODE As Long = 234
    Const LEFT_CODE As Long = 231
    Const RIGHT_CODE As Long = 232
    Const BODY_SEGMENT As Long = 110
    Const MOUSE As Long = 56
    Const MOUSE_HIGHLIGHT As Long = 65535
    Const DELIMITER As String = ","
    Const START As String = "16,16"
    Const START_PATH As String = "$P$16"
    Const LEGAL As Long = 1
    Const ILLEGAL As Long = 0
    Const GAME_MIN_CELLS_VALUE As Long = 2
    Const GAME_MAX_CELLS_VALUE As Long = 31
    Const FREEZE_PANE_PIVOT As Long = 40
    Const GAME_ZOOM As Long = 100
    Const TIME_ITERATION_VALUE As String = "00:00:01"
    Const MAXIMUM_RIBBON_HEIGHT As Long = 70

    Public timerActive As Boolean

Public Sub Main(ByVal currentLocation As Range, ByVal targetLocation As Range)
    Dim snakeString As String
    snakeString = Range("PathString").value
    Dim snakePath() As Range
    GetRangesFromString snakePath(), snakeString
    Dim isLegal As Boolean
    isLegal = True
    Dim verticalMovement As Long
    Dim horizontalMovement As Long
    horizontalMovement = CalculateMovement(targetLocation.Column, currentLocation.Column)
    verticalMovement = CalculateMovement(targetLocation.Row, currentLocation.Row)

    isLegal = CheckLegal(verticalMovement, horizontalMovement)
    If Not isLegal Then
        Range("LegalMove") = ILLEGAL
        currentLocation.Select
        Exit Sub
    End If
    
    Dim canMove As Boolean
    canMove = False
    If Not IsEmpty(targetLocation) Then
        canMove = CanContinue(targetLocation)
        If Not canMove Then
            Stop_Timing
            MsgBox "SCORE: " & UBound(snakePath)
            ResetBoard
            Exit Sub
        End If
        targetLocation.Interior.Color = xlNone
        PlaceMouse
    End If
    
    DrawSnakeHead targetLocation, horizontalMovement, verticalMovement
    Range("Position") = targetLocation.Row & DELIMITER & targetLocation.Column
    
    If UBound(snakePath) > 0 Then currentLocation.value = Chr$(BODY_SEGMENT)

    Range("HorizontalMovement").value = horizontalMovement
    Range("VerticalMovement").value = verticalMovement
    
    If canMove Then
        ReDim Preserve snakePath(LBound(snakePath) To UBound(snakePath) + 1)
    Else
        redraw snakePath()
    End If
    
    Set snakePath(UBound(snakePath)) = targetLocation
    snakeString = WritePath(snakePath)
    Range("PathString") = Replace(snakeString, "$", vbNullString)
    
    
End Sub
    
Private Sub GetRangesFromString(ByRef snakePath() As Range, ByVal snakeString As String)
    Dim snakePathString As Variant
    snakePathString = Split(snakeString, DELIMITER)
    ReDim snakePath(LBound(snakePathString) To UBound(snakePathString))
    Dim index As Long
    For index = LBound(snakePathString) To UBound(snakePathString)
        Set snakePath(index) = Range(snakePathString(index))
    Next
End Sub

Private Function CalculateMovement(ByVal ending As Long, ByVal beginning As Long) As Long
    If ending > beginning Then
        CalculateMovement = 1
    ElseIf beginning > ending Then
        CalculateMovement = -1
    Else
        CalculateMovement = 0
    End If
End Function

Private Function CheckLegal(ByVal verticalMovement As Long, ByVal horizontalMovement As Long) As Boolean
    If horizontalMovement = 0 Then
        If verticalMovement + Range("VerticalMovement") = 0 Then
            CheckLegal = ILLEGAL
            Exit Function
        Else
            CheckLegal = LEGAL
        End If
    ElseIf verticalMovement = 0 Then
        If horizontalMovement + Range("HorizontalMovement") = 0 Then
            CheckLegal = ILLEGAL
            Exit Function
        Else
            CheckLegal = LEGAL
        End If
    End If
End Function

Private Function CanContinue(ByVal targetLocation As Range) As Boolean
    If InStr(1, targetLocation.value, Chr$(BODY_SEGMENT)) > 0 Then
        CanContinue = False
    Else
        CanContinue = True
    End If
End Function

Private Sub DrawSnakeHead(ByVal targetLocation As Range, ByVal horizontalMovement As Long, ByVal verticalMovement As Long)
    Dim head As Long
    If horizontalMovement = 0 Then
        If verticalMovement = -1 Then
            head = UP_CODE
        Else
            head = DOWN_CODE
        End If
    Else
        If horizontalMovement = 1 Then
            head = RIGHT_CODE
        Else
            head = LEFT_CODE
        End If
    End If
    targetLocation.value = Chr$(head)
End Sub

Private Sub redraw(ByRef snakePath() As Range)
    Dim index As Long
    snakePath(LBound(snakePath)).ClearContents
    For index = LBound(snakePath) To UBound(snakePath) - 1
        Set snakePath(index) = snakePath(index + 1)
    Next
End Sub

Private Function WritePath(ByRef snakePath() As Range) As String
    Dim index As Long
    Dim tempString As String
    For index = LBound(snakePath) To UBound(snakePath)
        tempString = tempString & DELIMITER & snakePath(index).Address
    Next
    WritePath = Right$(tempString, Len(tempString) - 1)
End Function

Public Sub DrawGameBoard()
    Const SNAKE_FONT As String = "Wingdings"
    Const SNAKE_FONT_BOLD As Boolean = True
    Const SNAKE_FONT_SIZE As Long = 12
    Const COLUMN_WIDTH As Double = 3
    Const ROW_HEIGHT As Double = 21.75
    Dim borders As Range
    With GameSheet
        Dim boardRange As Range
        Dim gameRange As Range
        Set boardRange = .Range("A1:AF32")
        boardRange.Name = "Board"
        Set gameRange = .Range("B2:AD31")
        gameRange.Name = "GameRange"
        With boardRange
            .Clear
            .Font.Size = SNAKE_FONT_SIZE
            .Font.Name = SNAKE_FONT
            .Font.Bold = SNAKE_FONT_BOLD
            .Columns.ColumnWidth = COLUMN_WIDTH
            .Rows.RowHeight = ROW_HEIGHT
            .Rows(1).Name = "TopBorder"
            .Rows(100).EntireRow.Hidden = True
            .Rows(32).Name = "BottomBorder"
            .Columns(1).Name = "RightBorder"
            .Columns(32).Name = "LeftBorder"
            .Cells(100, 1).Name = "Position"
            .Cells(100, 2).Name = "PathString"
            .Cells(100, 3).Name = "FirstMove"
            .Cells(100, 4).Name = "HorizontalMovement"
            .Cells(100, 5).Name = "VerticalMovement"
            .Cells(100, 6).Name = "LegalMove"
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
        End With
        .Range("TopBorder").Interior.Color = vbBlack
        .Range("BottomBorder").Interior.Color = vbBlack
        .Range("RightBorder").Interior.Color = vbBlack
        .Range("LeftBorder").Interior.Color = vbBlack
        FreezeThePanes FREEZE_PANE_PIVOT, FREEZE_PANE_PIVOT
    End With
    Set borders = Application.Union(Range("TopBorder"), Range("BottomBorder"), Range("LeftBorder"), Range("RightBorder"))
    borders.Name = "Borders"
    For Each boardRange In Range("Borders")
        boardRange.value = Chr$(BODY_SEGMENT)
    Next
    If CommandBars("Ribbon").Height > MAXIMUM_RIBBON_HEIGHT Then CommandBars.ExecuteMso ("MinimizeRibbon")
    ActiveWindow.Zoom = GAME_ZOOM
    ResetBoard
End Sub
Private Sub ResetBoard()
    
    With GameSheet
        .Range("FirstMove") = 1
        .Range("HorizontalMovement") = 0
        .Range("VerticalMovement") = 0
        .Range("GameRange").ClearContents
        .Range("GameRange").Interior.Color = xlNone
        .Range("Position").value = START
        .Range("PathString").value = START_PATH
        .Range("LegalMove").value = LEGAL
        .Cells(16, 16) = Chr$(BODY_SEGMENT)
        .Cells(16, 16).Select
    End With
    
    PlaceMouse
    Stop_Timing
End Sub
Private Sub PlaceMouse()
    Dim randRow As Long
    Dim randColumn As Long
TryAgain:
    randRow = Int((GAME_MAX_CELLS_VALUE - GAME_MIN_CELLS_VALUE + 1) * Rnd + GAME_MIN_CELLS_VALUE)
    randColumn = Int((GAME_MAX_CELLS_VALUE - GAME_MIN_CELLS_VALUE + 1) * Rnd + GAME_MIN_CELLS_VALUE)
    If IsEmpty(GameSheet.Cells(randRow, randColumn)) Then
        GameSheet.Cells(randRow, randColumn).value = Chr$(MOUSE)
        GameSheet.Cells(randRow, randColumn).Interior.Color = MOUSE_HIGHLIGHT
    Else: GoTo TryAgain
    End If
End Sub

Private Sub FreezeThePanes(ByVal fRow As Long, ByVal fColumn As Long)
    With ActiveWindow
        .SplitColumn = fColumn
        .SplitRow = fRow
        .FreezePanes = True
    End With
End Sub

Public Sub Start_Timing()
    timerActive = True
    Application.OnTime Now + TimeValue(TIME_ITERATION_VALUE), "Timing"
End Sub

Public Sub Stop_Timing()
    timerActive = False
End Sub

Private Sub Timing()
    With GameSheet
        Dim repeatInterval As Date
        Dim horizontalMomentum As Long
        horizontalMomentum = Range("HorizontalMovement")
        Dim verticalMomentum As Long
        verticalMomentum = Range("VerticalMovement")
        If timerActive Then
            If horizontalMomentum = 0 Then
                MoveVertical verticalMomentum
            ElseIf verticalMomentum = 0 Then
                MoveHorizontal horizontalMomentum
            End If
        Else
            Exit Sub
        End If
        repeatInterval = Now + TimeValue(TIME_ITERATION_VALUE)
        Application.OnTime repeatInterval, "Timing"
    End With
End Sub

Private Sub MoveVertical(ByVal direction As Long)
    Dim timeTarget As Range
    Set timeTarget = Selection.Offset(direction)
    timeTarget.Select
End Sub

Private Sub MoveHorizontal(ByVal direction As Long)
    Dim timeTarget As Range
    Set timeTarget = Selection.Offset(, direction)
    timeTarget.Select
End Sub


