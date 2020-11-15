Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("King_NET.King")> Public Class King
    Inherits ChessPiece
    'King class allows for multiple instances of Chess King

    Public Overrides Function CheckMoves() As Byte
        'Add each legal move to the NextMove array of legal moves, 
        'Therefore CheckMoves must be called each time before moving using the nextmove array in a computer move
        'Each move has a score:
        ' 1 indicates a legal move 
        ' 2 score a possible taking move
        Dim posX As Short
        Dim posY As Short
        Dim posCount As Byte
        'Dim blocked As Boolean
        Dim result As Short

        posCount = 0
        If OnBoard Then
            posX = MyBase.XPos
            posY = MyBase.YPos + 1 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result <> Direction Then
                    posCount += 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    If result = -Direction Then 'opponent
                        MyBase.SetScore(posCount, 2) '2 for taking
                    Else
                        MyBase.SetScore(posCount, 1) '2 for taking
                    End If
                End If
            End If
            posX = MyBase.XPos + 1 'right
            posY = MyBase.YPos + 1 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result <> Direction Then
                    posCount += 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    If result = -Direction Then 'opponent
                        MyBase.SetScore(posCount, 2) '2 for taking
                    Else
                        MyBase.SetScore(posCount, 1) '2 for taking
                    End If
                End If
            End If
            posX = MyBase.XPos + 1 'right
            posY = MyBase.YPos
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result <> Direction Then
                    posCount += 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    If result = -Direction Then 'opponent
                        MyBase.SetScore(posCount, 2) '2 for taking
                    Else
                        MyBase.SetScore(posCount, 1) '2 for taking
                    End If
                End If
            End If
            posX = MyBase.XPos + 1 'right
            posY = MyBase.YPos - 1 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result <> Direction Then
                    posCount += 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    If result = -Direction Then 'opponent
                        MyBase.SetScore(posCount, 2) '2 for taking
                    Else
                        MyBase.SetScore(posCount, 1) '2 for taking
                    End If
                End If
            End If
            posX = MyBase.XPos
            posY = MyBase.YPos - 1 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result <> Direction Then
                    posCount += 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    If result = -Direction Then 'opponent
                        MyBase.SetScore(posCount, 2) '2 for taking
                    Else
                        MyBase.SetScore(posCount, 1) '2 for taking
                    End If
                End If
            End If
            posX = MyBase.XPos - 1 'left
            posY = MyBase.YPos - 1 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result <> Direction Then
                    posCount += 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    If result = -Direction Then 'opponent
                        MyBase.SetScore(posCount, 2) '2 for taking
                    Else
                        MyBase.SetScore(posCount, 1) '2 for taking
                    End If
                End If
            End If
            posX = MyBase.XPos - 1 'left
            posY = MyBase.YPos
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result <> Direction Then
                    posCount += 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    If result = -Direction Then 'opponent
                        MyBase.SetScore(posCount, 2) '2 for taking
                    Else
                        MyBase.SetScore(posCount, 1) '2 for taking
                    End If
                End If
            End If
            posX = MyBase.XPos - 1 'left
            posY = MyBase.YPos + 1 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result <> Direction Then
                    posCount += 1
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
        MyBase.MaxMove = posCount
        CheckMoves = posCount
    End Function

    Public Overrides Function CheckSupport() As Byte
        'Add each support provided to the mSupport array (analogous to mSupport array)
        Dim posX As Short
        Dim posY As Short
        Dim posCount As Byte 'index into mSupport Array
        'Dim blocked As Boolean 'only applies to queen, rook, bishop
        Dim result As Short

        posCount = 0 'for each Support possibility: index into mSupport Array
        If OnBoard Then
            posX = MyBase.XPos
            posY = MyBase.YPos + 1 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then 'own side Piece is there; add to mSupport array
                    posCount += 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
            posX = MyBase.XPos + 1 'right
            posY = MyBase.YPos + 1 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then
                    posCount += 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
            posX = MyBase.XPos + 1 'right
            posY = MyBase.YPos
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then
                    posCount += 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
            posX = MyBase.XPos + 1 'right
            posY = MyBase.YPos - 1 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then
                    posCount += 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
            posX = MyBase.XPos
            posY = MyBase.YPos - 1 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then
                    posCount += 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
            posX = MyBase.XPos - 1 'left
            posY = MyBase.YPos - 1 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then
                    posCount += 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
            posX = MyBase.XPos - 1 'left
            posY = MyBase.YPos
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then
                    posCount += 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
            posX = MyBase.XPos - 1 'left
            posY = MyBase.YPos + 1 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then
                    posCount += 1
                    MyBase.SetSuppX(posCount, posX)
                    MyBase.SetSuppY(posCount, posY)
                End If
            End If
        End If
        CheckSupport = posCount
    End Function

    Public Sub New()
        MyBase.New("K", 10)
    End Sub
	
End Class