Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Bishop_NET.Bishop")> Public Class Bishop
    Inherits ChessPiece

    Public Overrides Function CheckMoves() As Byte
        'Add each legal move to the NextMove array of legal moves, 
        'Therefore CheckMoves must be called each time before moving using the nextmove array in a computer move
        'Each move has a score:
        ' 1 indicates a legal move 
        ' 2 score a possible taking move
        Dim posX As Short
        Dim posY As Short
        Dim posCount As Byte
        Dim blocked As Boolean
        Dim result As Short

        posCount = 0 'first pos is 1
        If OnBoard Then
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            Do While Not blocked 'check each possibility in this direction
                posX += 1 'right
                posY += 1 'up
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction) '0=empty, direction=same side, -direction=opponent
                    If result = 0 Then 'the destination is empty
                        posCount += 1
                        MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                        MyBase.SetMoveY(posCount, posY)
                        MyBase.SetScore(posCount, 1) '1 for move to blank
                    Else 'the destination is occupied
                        If result = -Direction Then 'can take
                            posCount += 1
                            MyBase.SetMoveX(posCount, posX)
                            MyBase.SetMoveY(posCount, posY)
                            MyBase.SetScore(posCount, 2) '2 for taking
                        Else 'in terms of support the following is possible but we're not using it
                            'if 1 <= x <= 8 and 1 <= y <= 8, the bishop is currently supporting another Piece
                        End If
                        blocked = True 'but it's the end of line
                    End If
                Else
                    blocked = True
                End If
            Loop
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            Do While Not blocked
                posX += 1 'right
                posY -= 1 'down
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction)
                    If result = 0 Then
                        posCount += 1
                        MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                        MyBase.SetMoveY(posCount, posY)
                        MyBase.SetScore(posCount, 1) '1 for move to blank
                    Else
                        If result = -Direction Then 'can take
                            posCount += 1
                            MyBase.SetMoveX(posCount, posX)
                            MyBase.SetMoveY(posCount, posY)
                            MyBase.SetScore(posCount, 2) '2 for taking
                        End If
                        blocked = True
                    End If
                Else
                    blocked = True
                End If
            Loop
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            Do While Not blocked 'find all possible moves in this direction
                posX -= 1 'left
                posY -= 1 'down
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction) '0=empty, direction (own player or offboard), -direction (opponent)
                    If result = 0 Then 'empty
                        posCount += 1
                        MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                        MyBase.SetMoveY(posCount, posY)
                        MyBase.SetScore(posCount, 1) '1 for move to blank
                    Else
                        If result = -Direction Then 'opponent: can take
                            posCount += 1
                            MyBase.SetMoveX(posCount, posX)
                            MyBase.SetMoveY(posCount, posY)
                            MyBase.SetScore(posCount, 2) '2 for taking
                        End If
                        blocked = True 'no more possible moves in that direction
                    End If
                Else
                    blocked = True
                End If
            Loop
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            Do While Not blocked
                posX -= 1 'left
                posY += 1 'up
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction)
                    If result = 0 Then
                        posCount += 1
                        MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                        MyBase.SetMoveY(posCount, posY)
                        MyBase.SetScore(posCount, 1) '1 for move to blank
                    Else
                        If result = -Direction Then 'can take
                            posCount += 1
                            MyBase.SetMoveX(posCount, posX)
                            MyBase.SetMoveY(posCount, posY)
                            MyBase.SetScore(posCount, 2) '2 for taking
                        End If
                        blocked = True
                    End If
                Else
                    blocked = True
                End If
            Loop
        End If
        MyBase.MaxMove = posCount
        CheckMoves = posCount
    End Function

    Public Overrides Function CheckSupport() As Byte
        'Add each support provided to the mSupport array (analogous to nextmove array)
        Dim posX As Short 'current x co-ordinate to look at
        Dim posY As Short 'current y co-ordinate to look at
        Dim posCount As Byte 'each possible support: there are four max
        Dim blocked As Boolean 'examine each position till you can't go any further...
        Dim result As Short 'either 0 for nothing there, direction the same means it's your own Piece (i.e. supporting)

        posCount = 0 'first possibility is 1 in the array i.e. 1..4
        If OnBoard Then 'only has things it can support if this Piece is on the board
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            Do While Not blocked
                posX += 1 'right
                posY += 1 'up
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction) ' 0=empty, direction same=supported, direction opposite can take
                    If result <> 0 Then 'the destination is occupied
                        If result = Direction Then 'one of our own side found: therefore this Piece supports it
                            posCount += 1
                            MyBase.SetSuppX(posCount, posX)
                            MyBase.SetSuppY(posCount, posY)
                        End If
                        blocked = True 'can only directly support one in any direction
                    End If
                Else
                    blocked = True 'off the board
                End If
            Loop
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            result = 0
            Do While Not blocked
                posX += 1 'right
                posY -= 1 'down
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    If Not blocked Then result = IsLegal(posX, posY, Direction)
                    If result <> 0 Then 'the destination is occupied
                        If result = Direction Then 'one of our own side found
                            posCount += 1
                            MyBase.SetSuppX(posCount, posX)
                            MyBase.SetSuppY(posCount, posY)
                        End If
                        blocked = True
                    End If
                Else 'reached the end of the board.
                    blocked = True
                End If
            Loop
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            result = 0
            Do While Not blocked
                If posX < 1 Then blocked = True
                If posY < 1 Then blocked = True
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    If Not blocked Then result = IsLegal(posX, posY, Direction)
                    If result <> 0 Then 'the destination is occupied
                        If result = Direction Then 'one of our own side found
                            posCount += 1
                            MyBase.SetSuppX(posCount, posX)
                            MyBase.SetSuppY(posCount, posY)
                        End If
                        blocked = True
                    End If
                Else
                    blocked = True
                End If
            Loop
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            result = 0
            Do While Not blocked
                posX -= 1 'left
                posY += 1 'up
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    If Not blocked Then result = IsLegal(posX, posY, Direction)
                    If result <> 0 Then 'the destination is occupied
                        If result = Direction Then 'one of our own side found
                            posCount += 1
                            MyBase.SetSuppX(posCount, posX)
                            MyBase.SetSuppY(posCount, posY)
                        End If
                        blocked = True
                    End If
                Else
                    blocked = True
                End If
            Loop
        End If
        MyBase.MaxMove = posCount
        CheckSupport = posCount
    End Function

    Public Sub New()
        MyBase.New("B", 4)
    End Sub

End Class