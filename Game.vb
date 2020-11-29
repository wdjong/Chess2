Public Class Game
    Const MAXMOVE As Integer = 200
    Private ReadOnly mMoves(MAXMOVE) As Move 'x= moves, y1=Pieceid, y2=xorig y3=yorig y4=xdest y5=ydest y6=Piece2id, y7 Piece2xorig y8=Piece2yorig y9=Piece2xdest y10=Piece2ydest
    Private mCount As Short
    Private mPlaying As Boolean 'keep track of whether we're playing or not at the moment
    Public ReadOnly Property Count() As Integer
        'returns how many moves have been made
        Get
            Count = mCount
        End Get
    End Property

    Public ReadOnly Property Moves(ByVal i As Integer) As Move
        'return a specific move
        Get
            Moves = mMoves(i)
        End Get
    End Property

    Public Property Playing() As Boolean
        'keep track of whether we're playing or not at the moment
        Get
            Playing = mPlaying
        End Get
        Set(ByVal value As Boolean)
            mPlaying = value
        End Set
    End Property

    Public Sub Add(ByVal aMove As Move)
        'Create a new move object and add a reference to it to the list
        Dim newMove As New Move With {
            .P1id = aMove.P1id,
            .P1XOrig = aMove.P1XOrig,
            .P1YOrig = aMove.P1YOrig,
            .P1XDest = aMove.P1XDest,
            .P1YDest = aMove.P1YDest,
            .P2id = aMove.P2id,
            .P2XOrig = aMove.P2XOrig,
            .P2YOrig = aMove.P2YOrig,
            .P2XDest = aMove.P2XDest,
            .P2YDest = aMove.P2YDest,
            .MCode = aMove.MCode
        }
        If mCount > MAXMOVE - 1 Then
            Cout("Stalemate: Draw (> max move)")
            'MessageBox.Show(aWhite.GetRawScore())
            'SaveResult(aWhite.GetRawScore() + 20)
            aGame.Playing = False
        Else
            mMoves(mCount) = newMove 'first move is 0
            mCount += 1 'mCount is 1
        End If
    End Sub

    Public Sub Remove()
        'Destroy the last move
        Dim aMove As Move

        mCount -= 1
        aMove = mMoves(mCount)
    End Sub

    Public Function Last() As Move
        'return the last move
        Try
            Last = mMoves(mCount - 1)
        Catch ex As Exception
            MsgBox("Can't undo")
            Last = Nothing
        End Try
    End Function

    Public Sub New()
        mCount = 0
    End Sub

    Friend Sub Clear()
        mCount = 0
        Array.Clear(mMoves, 0, mMoves.Length)
    End Sub
End Class
