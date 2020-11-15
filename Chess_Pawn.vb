Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Pawn_NET.Pawn")> Public Class Pawn
    Inherits ChessPiece

    'Pawn class allows for multiple instances of Chess Pawn
    Const DIRUP As Short = 1
	Const DIRDOWN As Short = -1

    Public Sub New()
        MyBase.New("P", 2)
    End Sub

    Public Overrides Function CheckMoves() As Byte
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
            posX = MyBase.XPos
            posY = MyBase.YPos + Direction 'a pawn can move forward 1 rank
            result = IsLegal(posX, posY, Direction) '0 = Legal (?!) (returns team direction (1/-1) if illegal)
            If result = 0 Then 'Unoccupied: must be unoccupied for a pawn to move forward
                posCount += 1 'increment legal move count
                MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                MyBase.SetMoveY(posCount, posY)
                MyBase.SetScore(posCount, 1) '1 for move to blank
                If (MyBase.YPos = 2 And Direction = DIRUP) Or (MyBase.YPos = 7 And DIRDOWN) Then 'on 2nd rank; might be able to move forward 2
                    posY = MyBase.YPos + 2 * Direction
                    result = IsLegal(posX, posY, Direction)
                    If result = 0 Then 'Unoccupied: must be unoccupied for a pawn to move forward
                        posCount += 1 'increment legal move count
                        MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                        MyBase.SetMoveY(posCount, posY)
                        MyBase.SetScore(posCount, 1) '1 for move to blank
                    End If
                End If
            End If
            'check for taking Pieces (pawn can take diagonally)
            posX = MyBase.XPos + 1 'right
            posY = MyBase.YPos + Direction 'forward
            result = IsLegal(posX, posY, Direction) '0 = blank, -Direction = opponent, Direction = off board
            If result = -Direction Then 'Opponent found
                posCount += 1
                MyBase.SetMoveX(posCount, posX)
                MyBase.SetMoveY(posCount, posY)
                MyBase.SetScore(posCount, 2) '2 for taking
            End If
            posX = MyBase.XPos - 1 'left
            posY = MyBase.YPos + Direction 'forward
            result = IsLegal(posX, posY, Direction)
            If result = -Direction Then 'can take
                posCount += 1
                MyBase.SetMoveX(posCount, posX)
                MyBase.SetMoveY(posCount, posY)
                MyBase.SetScore(posCount, 2) '2 for taking
            End If
        End If
        MyBase.MaxMove = posCount
        CheckMoves = posCount
    End Function

    Public Overrides Function CheckSupport() As Byte
        'Add each support provided to the mSupport array (analogous to nextmove array)
        Dim posX As Short 'the destination
        Dim posY As Short 'the destination co-ord
        Dim posCount As Byte 'the count of possible moves
        Dim result As Short

        posCount = 0
        If OnBoard Then
            'check for taking Pieces
            posX = MyBase.XPos + 1 'right
            posY = MyBase.YPos + Direction 'forward
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then 'there's someone on our side to support
                    posCount += 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
            posX = MyBase.XPos - 1 'left
            posY = MyBase.YPos + Direction 'forward
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then 'can take
                    posCount += 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
        End If
        CheckSupport = posCount
    End Function

End Class