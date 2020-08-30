Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Rook_NET.Rook")> Public Class Rook
    Inherits ChessPiece

    'Rook class allows for multiple instances of Chess Rook
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
        Dim blocked As Boolean
        Dim result As Short

        posCount = 0
        If OnBoard Then
            blocked = False
            posX = MyBase.xpos
            posY = MyBase.ypos
            Do While Not blocked
                posX = posX + 1 'right
                result = isLegal(posX, posY, Direction)
                If result = 0 Then 'nothing encountered
                    posCount = posCount + 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    MyBase.SetScore(posCount, 1) '1 for move to blank
                Else 'It can take this piece
                    If result = -Direction Then 'can take
                        posCount = posCount + 1
                        MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                        MyBase.SetMoveY(posCount, posY)
                        MyBase.SetScore(posCount, 2) '1 for move to blank
                    End If
                    blocked = True
                End If
            Loop
            blocked = False
            posX = MyBase.xpos
            posY = MyBase.ypos
            Do While Not blocked
                posY = posY - 1 'down
                result = isLegal(posX, posY, Direction)
                If result = 0 Then
                    posCount = posCount + 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    MyBase.SetScore(posCount, 1) '1 for move to blank
                Else
                    If result = -Direction Then 'can take
                        posCount = posCount + 1
                        MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                        MyBase.SetMoveY(posCount, posY)
                        MyBase.SetScore(posCount, 2) '1 for move to blank
                    End If
                    blocked = True
                End If
            Loop
            blocked = False
            posX = MyBase.xpos
            posY = MyBase.ypos
            Do While Not blocked
                posX = posX - 1 'left
                result = isLegal(posX, posY, Direction)
                If result = 0 Then
                    posCount = posCount + 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    MyBase.SetScore(posCount, 1) '1 for move to blank
                Else
                    If result = -Direction Then 'can take
                        posCount = posCount + 1
                        MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                        MyBase.SetMoveY(posCount, posY)
                        MyBase.SetScore(posCount, 2) '1 for move to blank
                    End If
                    blocked = True
                End If
            Loop
            blocked = False
            posX = MyBase.xpos
            posY = MyBase.ypos
            Do While Not blocked
                posY = posY + 1 'up
                result = isLegal(posX, posY, Direction)
                If result = 0 Then
                    posCount = posCount + 1
                    MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                    MyBase.SetMoveY(posCount, posY)
                    MyBase.SetScore(posCount, 1) '1 for move to blank
                Else
                    If result = -Direction Then 'can take
                        posCount = posCount + 1
                        MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                        MyBase.SetMoveY(posCount, posY)
                        MyBase.SetScore(posCount, 2) '1 for move to blank
                    End If
                    blocked = True
                End If
            Loop
        End If
        MyBase.MaxMove = posCount
        checkMoves = posCount
    End Function

    Public Overrides Function checkSupport() As Byte
        'Add each support provided to the mSupport array (analogous to mSupport array)
        Dim posX As Short
        Dim posY As Short
        Dim posCount As Byte
        Dim blocked As Boolean
        Dim result As Short

        posCount = 0
        If OnBoard Then
            blocked = False
            posX = MyBase.xPos
            posY = MyBase.yPos
            Do While Not blocked
                posX = posX + 1 'right
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = isLegal(posX, posY, Direction)
                    If result = Direction Then 'can take
                        posCount = posCount + 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                        blocked = True
                    End If
                Else
                    blocked = True
                End If
            Loop
            blocked = False
            posX = MyBase.xPos
            posY = MyBase.yPos
            Do While Not blocked
                posY = posY - 1 'down
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = isLegal(posX, posY, Direction)
                    If result = Direction Then 'can take
                        posCount = posCount + 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                        blocked = True
                    End If
                Else
                    blocked = True
                End If
            Loop
            blocked = False
            posX = MyBase.xPos
            posY = MyBase.yPos
            Do While Not blocked
                posX = posX - 1 'left
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = isLegal(posX, posY, Direction)
                    If result = Direction Then 'can take
                        posCount = posCount + 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                        blocked = True
                    End If
                Else
                    blocked = True
                End If
            Loop
            blocked = False
            posX = MyBase.xPos
            posY = MyBase.yPos
            Do While Not blocked
                posY = posY + 1 'up
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = isLegal(posX, posY, Direction)
                    If result = Direction Then 'can take
                        posCount = posCount + 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                        blocked = True
                    End If
                Else
                    blocked = True
                End If
            Loop
        End If
        checkSupport = posCount
    End Function

    Public Sub New()
        MyBase.New("R", 6)
    End Sub

End Class