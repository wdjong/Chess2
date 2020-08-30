Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Knight_NET.Knight")> Public Class Knight
    Inherits ChessPiece   'Knight class allows for multiple instances of Chess Knights

    Const DIRUP As Short = 1
    Const DIRDOWN As Short = -1
    Const XMOVE As Short = 0
	Const YMOVE As Short = 1
	Const SCOREMOVE As Short = 2
    Const SUPPVALUE As Byte = 2

	
    Public Overrides Function checkMoves() As Byte
        'Add each legal move to the NextMove array of legal moves, 
        'Therefore CheckMoves must be called each time before moving using the nextmove array in a computer move
        'Each move has a score:
        ' 1 indicates a legal move 
        ' 2 score a possible taking move
        Dim posX As Short
        Dim posY As Short
        Dim posCount As Byte
        Dim result As Short

        posCount = 0
        If OnBoard Then 'Knight is in play
            posX = MyBase.xPos + 1 'right (going clockwise)
            posY = MyBase.yPos + 2 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result <> Direction Then 'empty or opposing piece
                    posCount = posCount + 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    If result = -Direction Then 'opponent
                        MyBase.SetScore(posCount, 2) '2 for taking
                    Else
                        MyBase.SetScore(posCount, 1) '2 for taking
                    End If
                End If
            End If
            posX = MyBase.xPos + 2 'right
            posY = MyBase.yPos + 1 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result <> Direction Then
                    posCount = posCount + 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    If result = -Direction Then 'opponent
                        MyBase.SetScore(posCount, 2) '2 for taking
                    Else
                        MyBase.SetScore(posCount, 1) '2 for taking
                    End If
                End If
            End If
            posX = MyBase.xPos + 2 'right
            posY = MyBase.yPos - 1 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result <> Direction Then
                    posCount = posCount + 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    If result = -Direction Then 'opponent
                        MyBase.SetScore(posCount, 2) '2 for taking
                    Else
                        MyBase.SetScore(posCount, 1) '2 for taking
                    End If
                End If
            End If
            posX = MyBase.xPos + 1 'right
            posY = MyBase.yPos - 2 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result <> Direction Then
                    posCount = posCount + 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    If result = -Direction Then 'opponent
                        MyBase.SetScore(posCount, 2) '2 for taking
                    Else
                        MyBase.SetScore(posCount, 1) '2 for taking
                    End If
                End If
            End If
            posX = MyBase.xPos - 1 'left
            posY = MyBase.yPos - 2 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result <> Direction Then
                    posCount = posCount + 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    If result = -Direction Then 'opponent
                        MyBase.SetScore(posCount, 2) '2 for taking
                    Else
                        MyBase.SetScore(posCount, 1) '2 for taking
                    End If
                End If
            End If
            posX = MyBase.xPos - 2 'left
            posY = MyBase.yPos - 1 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result <> Direction Then
                    posCount = posCount + 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    If result = -Direction Then 'opponent
                        MyBase.SetScore(posCount, 2) '2 for taking
                    Else
                        MyBase.SetScore(posCount, 1) '2 for taking
                    End If
                End If
            End If
            posX = MyBase.xPos - 2 'left
            posY = MyBase.yPos + 1 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result <> Direction Then
                    posCount = posCount + 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    If result = -Direction Then 'opponent
                        MyBase.SetScore(posCount, 2) '2 for taking
                    Else
                        MyBase.SetScore(posCount, 1) '2 for taking
                    End If
                End If
            End If
            posX = MyBase.xPos - 1 'left
            posY = MyBase.yPos + 2 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result <> Direction Then
                    posCount = posCount + 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    If result = -Direction Then 'opponent
                        MyBase.SetScore(posCount, 2) '2 for taking
                    Else
                        MyBase.SetScore(posCount, 1) '2 for taking
                    End If
                End If
            End If
        End If
        MyBase.MaxMove = posCount 'Number of valid moves 1 to n
        checkMoves = posCount
    End Function

    Public Overrides Function checkSupport() As Byte
        'Add each support provided to the mSupport array (analogous to mSupport array)
        Dim posX As Short
        Dim posY As Short
        Dim posCount As Byte
        Dim result As Short

        posCount = 0
        If OnBoard Then 'Knight is in play
            posX = MyBase.xPos + 1 'right (going clockwise)
            posY = MyBase.yPos + 2 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result = Direction Then 'empty or opposing piece
                    posCount = posCount + 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
            posX = MyBase.xPos + 2 'right
            posY = MyBase.yPos + 1 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result = Direction Then
                    posCount = posCount + 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
            posX = MyBase.xPos + 2 'right
            posY = MyBase.yPos - 1 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result = Direction Then
                    posCount = posCount + 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
            posX = MyBase.xPos + 1 'right
            posY = MyBase.yPos - 2 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result = Direction Then
                    posCount = posCount + 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
            posX = MyBase.xPos - 1 'left
            posY = MyBase.yPos - 2 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result = Direction Then
                    posCount = posCount + 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
            posX = MyBase.xPos - 2 'left
            posY = MyBase.yPos - 1 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result = Direction Then
                    posCount = posCount + 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
            posX = MyBase.xPos - 2 'left
            posY = MyBase.yPos + 1 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result = Direction Then
                    posCount = posCount + 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
            posX = MyBase.xPos - 1 'left
            posY = MyBase.yPos + 2 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result = Direction Then
                    posCount = posCount + 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
        End If
        checkSupport = posCount
    End Function

    Public Sub New()
        MyBase.New("N", 4)
    End Sub

End Class