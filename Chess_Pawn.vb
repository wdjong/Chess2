Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Pawn_NET.Pawn")> Public Class Pawn
    Inherits ChessPiece

    'Pawn class allows for multiple instances of Chess Pawn
    Const DIRUP As Short = 1
	Const DIRDOWN As Short = -1
    Const XMOVE As Short = 0
	Const YMOVE As Short = 1
	Const SCOREMOVE As Short = 2
    Const SUPPVALUE As Byte = 2

    Public Sub New()
        MyBase.New("P", 2)
    End Sub

    Public Overrides Function checkMoves() As Byte
        'Find the number of valid moves
        'Add each legal move to the NextMove array of legal moves, 
        'Therefore, this routine must be called each time before moving.
        'Each move has a score:
        ' 1 indicates a legal move 
        ' 2 score a possible taking move
        Dim posX As Short 'the destination
        Dim posY As Short 'the destination co-ord
        Dim posCount As Byte 'the count of possible moves
        Dim result As Short

        posCount = 0 'count of valid positions
        If OnBoard Then
            posX = MyBase.xPos
            posY = MyBase.yPos + Direction 'a pawn can move forward 1 rank
            result = isLegal(posX, posY, Direction) '0 = Legal (?!) (returns team direction (1/-1) if illegal)
            If result = 0 Then 'Unoccupied: must be unoccupied for a pawn to move forward
                posCount = posCount + 1 'increment legal move count
                MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                MyBase.SetMoveY(posCount, posY)
                MyBase.SetScore(posCount, 1) '1 for move to blank
                If (MyBase.yPos = 2 And Direction = DIRUP) Or (MyBase.yPos = 7 And DIRDOWN) Then 'on 2nd rank; might be able to move forward 2
                    posY = MyBase.yPos + 2 * Direction
                    result = isLegal(posX, posY, Direction)
                    If result = 0 Then 'Unoccupied: must be unoccupied for a pawn to move forward
                        posCount = posCount + 1 'increment legal move count
                        MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                        MyBase.SetMoveY(posCount, posY)
                        MyBase.SetScore(posCount, 1) '1 for move to blank
                    End If
                End If
            End If
            'check for taking pieces (pawn can take diagonally)
            posX = MyBase.xPos + 1 'right
            posY = MyBase.yPos + Direction 'forward
            result = isLegal(posX, posY, Direction) '0 = blank, -Direction = opponent, Direction = off board
            If result = -Direction Then 'Opponent found
                posCount = posCount + 1
                MyBase.SetMoveX(posCount, posX)
                MyBase.SetMoveY(posCount, posY)
                MyBase.SetScore(posCount, 2) '2 for taking
            End If
            posX = MyBase.xPos - 1 'left
            posY = MyBase.yPos + Direction 'forward
            result = isLegal(posX, posY, Direction)
            If result = -Direction Then 'can take
                posCount = posCount + 1
                MyBase.SetMoveX(posCount, posX)
                MyBase.SetMoveY(posCount, posY)
                MyBase.SetScore(posCount, 2) '2 for taking
            End If
        End If
        MyBase.MaxMove = posCount
        checkMoves = posCount
    End Function

    Public Overrides Function checkSupport() As Byte
        'Add each support provided to the mSupport array (analogous to nextmove array)
        Dim posX As Short 'the destination
        Dim posY As Short 'the destination co-ord
        Dim posCount As Byte 'the count of possible moves
        Dim result As Short

        posCount = 0
        If OnBoard Then
            'check for taking pieces
            posX = MyBase.xPos + 1 'right
            posY = MyBase.yPos + Direction 'forward
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result = Direction Then 'there's someone on our side to support
                    posCount = posCount + 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
            posX = MyBase.xPos - 1 'left
            posY = MyBase.yPos + Direction 'forward
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = isLegal(posX, posY, Direction)
                If result = Direction Then 'can take
                    posCount = posCount + 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
        End If
        checkSupport = posCount
    End Function

End Class