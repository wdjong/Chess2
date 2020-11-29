Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Queen_NET.Queen")> Public Class Queen
    Inherits ChessPiece   'Queen class allows for multiple instances of Chess Queen

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

        posCount = 0
        If OnBoard Then
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            Do While Not blocked
                posY += 1 'up
                result = IsLegal(posX, posY, Direction) 'result 0 = empty, Same direction off board or own player, opposite direction opponent
                blocked = RecordResult(result, posCount, posX, posY)
            Loop
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            Do While Not blocked
                posX += 1 'right
                posY += 1 'up
                result = IsLegal(posX, posY, Direction)
                blocked = RecordResult(result, posCount, posX, posY)
            Loop
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            Do While Not blocked
                posX += 1 'right
                result = IsLegal(posX, posY, Direction)
                blocked = RecordResult(result, posCount, posX, posY)
            Loop
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            Do While Not blocked
                posX += 1 'right
                posX -= 1 'down
                result = IsLegal(posX, posY, Direction)
                blocked = RecordResult(result, posCount, posX, posY)
            Loop
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            Do While Not blocked
                posX -= 1 'down
                result = IsLegal(posX, posY, Direction)
                blocked = RecordResult(result, posCount, posX, posY)
            Loop
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            Do While Not blocked
                posX -= 1 'left
                posX -= 1 'up
                result = IsLegal(posX, posY, Direction)
                blocked = RecordResult(result, posCount, posX, posY)
            Loop
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            Do While Not blocked
                posX -= 1 'left
                result = IsLegal(posX, posY, Direction)
                blocked = RecordResult(result, posCount, posX, posY)
            Loop
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            Do While Not blocked
                posX -= 1 'left
                posY += 1 'up
                result = IsLegal(posX, posY, Direction)
                blocked = RecordResult(result, posCount, posX, posY)
            Loop
        End If
        MyBase.MaxMove = posCount
        CheckMoves = posCount
    End Function

    Private Function RecordResult(result As Short, posCount As Byte, posX As Short, posY As Short) As Boolean
        RecordResult = False 'not blocked
        If result = 0 Then
            posCount += 1
            If posCount >= MyBase.MAXPOSCOUNT Then
                Stop
            End If
            MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
            MyBase.SetMoveY(posCount, posY)
            MyBase.SetScore(posCount, 1) 'give it a modest positive score
        Else
            If result = -Direction Then 'can take
                posCount += 1
                MyBase.SetMoveX(posCount, posX) 'remember the co-ordinates
                MyBase.SetMoveY(posCount, posY)
                MyBase.SetScore(posCount, 2) 'give it a modest positive score
            End If
            RecordResult = True 'blocked
        End If
    End Function

    Public Overrides Function CheckSupport() As Byte
        'Add each support provided to the mSupport array (analogous to nextmove array)
        Dim posX As Short
        Dim posY As Short
        Dim posCount As Byte
        Dim blocked As Boolean
        Dim result As Short

        posCount = 0
        If OnBoard Then
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            Do While Not blocked
                posY += 1 'up
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction)
                    If result = Direction Then 'can support
                        posCount += 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                        blocked = True
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
                posY += 1 'up
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction)
                    If result = Direction Then 'can take
                        posCount += 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                        blocked = True
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
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction)
                    If result = Direction Then 'can take
                        posCount += 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                        blocked = True
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
                posX -= 1 'down
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction)
                    If result = Direction Then 'can take
                        posCount += 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                        blocked = True
                    End If
                Else
                    blocked = True
                End If
            Loop
            blocked = False
            posX = MyBase.XPos
            posY = MyBase.YPos
            Do While Not blocked
                posX -= 1 'down
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction)
                    If result = Direction Then 'can take
                        posCount += 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                        blocked = True
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
                posX -= 1 'up
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction)
                    If result = Direction Then 'can take
                        posCount += 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                        blocked = True
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
                If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                    result = IsLegal(posX, posY, Direction)
                    If result = Direction Then 'can take
                        posCount += 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                        blocked = True
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
                    If result = Direction Then 'can take
                        posCount += 1
                        MyBase.SetSuppX(posCount, posX)
                        MyBase.SetSuppY(posCount, posY)
                        blocked = True
                    End If
                Else
                    blocked = True
                End If
            Loop
        End If
        CheckSupport = posCount
    End Function

    Public Sub New()
        MyBase.New("Q", 8)
    End Sub
	
End Class