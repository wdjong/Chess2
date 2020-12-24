Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("King_NET.King")> Public Class King
    Inherits ChessPiece
    'King class allows for multiple instances of Chess King
    Public Sub New()
        MyBase.New("K", 10)
    End Sub

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
                    If Not MovingIntoCheck(posX, posY, Direction) Then
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
            posX = MyBase.XPos + 1 'right
            posY = MyBase.YPos + 1 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result <> Direction Then
                    If Not MovingIntoCheck(posX, posY, Direction) Then
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
            posX = MyBase.XPos + 1 'right
            posY = MyBase.YPos
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result <> Direction Then
                    If Not MovingIntoCheck(posX, posY, Direction) Then
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
            posX = MyBase.XPos + 1 'right
            posY = MyBase.YPos - 1 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result <> Direction Then
                    If Not MovingIntoCheck(posX, posY, Direction) Then

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
            posX = MyBase.XPos
            posY = MyBase.YPos - 1 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result <> Direction Then
                    If Not MovingIntoCheck(posX, posY, Direction) Then

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
            posX = MyBase.XPos - 1 'left
            posY = MyBase.YPos - 1 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result <> Direction Then
                    If Not MovingIntoCheck(posX, posY, Direction) Then

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
            posX = MyBase.XPos - 1 'left
            posY = MyBase.YPos
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result <> Direction Then
                    If Not MovingIntoCheck(posX, posY, Direction) Then

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

            posX = MyBase.XPos - 1 'left
            posY = MyBase.YPos + 1 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result <> Direction Then
                    If Not MovingIntoCheck(posX, posY, Direction) Then

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
        End If
        MyBase.MaxMove = posCount
        CheckMoves = posCount
    End Function

    Private Function MovingIntoCheck(posX As Short, posY As Short, direction As Short) As Boolean
        'direction is the colour of the king that could be moving into check i.e. 1 BottomGoingUp Red/White or -1 TopGoingDown Blue/Black
        'Check if any moves for opposition Pieces are threatening the King's proposed position
        Dim p As Byte 'each Piece
        Dim pf As Byte 'first Piece
        Dim pl As Byte 'last Piece

        MovingIntoCheck = False
        Try
            If direction = TOPGOINGDOWN Then 'blue: identify the range of opposition Pieces
                pf = 1 'look at red Pieces
                pl = 16
            Else 'BOTTOMGOINGUP
                pf = 17
                pl = MAXPIECE
            End If
            For p = pf To pl 'For each Piece in the opponent's team
                If aChessPiece(p).PotentialTake(posX, posY) Then
                    MovingIntoCheck = True
                End If
            Next p
        Catch ex As Exception
            MsgBox("MovingIntoCheck: " & ex.Message)
        End Try
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
                    If Not MovingIntoCheck(posX, posY, Direction) Then
                        posCount += 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                    End If
                End If
            End If
            posX = MyBase.XPos + 1 'right
            posY = MyBase.YPos + 1 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then
                    If Not MovingIntoCheck(posX, posY, Direction) Then
                        posCount += 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                    End If
                End If
            End If
            posX = MyBase.XPos + 1 'right
            posY = MyBase.YPos
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then
                    If Not MovingIntoCheck(posX, posY, Direction) Then
                        posCount += 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                    End If
                End If

            End If
            posX = MyBase.XPos + 1 'right
            posY = MyBase.YPos - 1 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then
                    If Not MovingIntoCheck(posX, posY, Direction) Then
                        posCount += 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                    End If
                End If
            End If
            posX = MyBase.XPos
            posY = MyBase.YPos - 1 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then
                    If Not MovingIntoCheck(posX, posY, Direction) Then
                        posCount += 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                    End If
                End If
            End If
            posX = MyBase.XPos - 1 'left
            posY = MyBase.YPos - 1 'down
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then
                    If Not MovingIntoCheck(posX, posY, Direction) Then
                        posCount += 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                    End If
                End If
            End If
            posX = MyBase.XPos - 1 'left
            posY = MyBase.YPos
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then
                    If Not MovingIntoCheck(posX, posY, Direction) Then
                        posCount += 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                    End If
                End If
            End If
            posX = MyBase.XPos - 1 'left
            posY = MyBase.YPos + 1 'up
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then
                    If Not MovingIntoCheck(posX, posY, Direction) Then
                        posCount += 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                    End If
                End If
            End If
        End If
        CheckSupport = posCount
    End Function

    Public Overrides Function PotentialTake(posTestX As Short, posTestY As Short) As Boolean
        'Check all possible moves of piece and if any is the move we're interested in then it's a legal move
        Dim posX As Short 'the destination
        Dim posY As Short 'the destination co-ord
        Dim result As Short
        PotentialTake = False
        Try
            If OnBoard Then
                posX = MyBase.XPos
                posY = MyBase.YPos
                posY += 1 'up
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction)   '0=empty, direction=same side, -direction=opponent (Note: doesn't check moving into/out of check)
                    If posX = posTestX And posY = posTestY Then
                        PotentialTake = True
                        Exit Function
                    End If
                End If
                posX = MyBase.XPos
                posY = MyBase.YPos
                posX += 1 'right
                posY += 1 'up
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction)   '0=empty, direction=same side, -direction=opponent
                    If posX = posTestX And posY = posTestY Then
                        PotentialTake = True
                        Exit Function
                    End If
                End If
                posX = MyBase.XPos
                posY = MyBase.YPos
                posX += 1 'right
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction)   '0=empty, direction=same side, -direction=opponent
                    If posX = posTestX And posY = posTestY Then
                        PotentialTake = True
                        Exit Function
                    End If
                End If
                posX = MyBase.XPos
                posY = MyBase.YPos
                posX += 1 'right
                posY -= 1 'down
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction)   '0=empty, direction=same side, -direction=opponent
                    If posX = posTestX And posY = posTestY Then
                        PotentialTake = True
                        Exit Function
                    End If
                End If
                posX = MyBase.XPos
                posY = MyBase.YPos
                posY -= 1 'down
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction)   '0=empty, direction=same side, -direction=opponent
                    If posX = posTestX And posY = posTestY Then
                        PotentialTake = True
                        Exit Function
                    End If
                End If
                posX = MyBase.XPos
                posY = MyBase.YPos
                posX -= 1 'left
                posY -= 1 'down
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction)   '0=empty, direction=same side, -direction=opponent
                    If posX = posTestX And posY = posTestY Then
                        PotentialTake = True
                        Exit Function
                    End If
                End If
                posX = MyBase.XPos
                posY = MyBase.YPos
                posX -= 1 'left
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction)   '0=empty, direction=same side, -direction=opponent
                    If posX = posTestX And posY = posTestY Then
                        PotentialTake = True
                        Exit Function
                    End If
                End If
                posX = MyBase.XPos
                posY = MyBase.YPos
                posX -= 1 'left
                posY += 1 'up
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction)   '0=empty, direction=same side, -direction=opponent
                    If posX = posTestX And posY = posTestY Then
                        PotentialTake = True
                        Exit Function
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox("PotentialTake K: " & ex.Message)
        End Try
    End Function


End Class