Public Class Game
    Const mMAXMOVES As Short = 200
    Private mMoves(mMAXMOVES) As Move 'x= moves, y1=pieceid, y2=xorig y3=yorig y4=xdest y5=ydest y6=piece2id, y7 piece2xorig y8=piece2yorig y9=piece2xdest y10=piece2ydest
    Private mCount As Short
    Private mPlaying As Boolean 'keep track of whether we're playing or not at the moment
    Private mShareThoughts As Boolean = False

    Public ReadOnly Property count() As Integer
        'returns how many moves have been made
        Get
            count = mCount
        End Get
    End Property

    Public Property ShareThoughts() As Boolean
        'whether we want to output thinking
        Get
            ShareThoughts = mShareThoughts
        End Get
        Set(ByVal value As Boolean)
            mShareThoughts = value
        End Set
    End Property

    Public ReadOnly Property moves(ByVal i As Integer) As Move
        'return a specific move
        Get
            moves = mMoves(i)
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

    Public Sub add(ByVal aMove As Move)
        'Create a new move object and add a reference to it to the list
        Dim newMove As New Move

        newMove.p1id = aMove.p1id
        newMove.p1XOrig = aMove.p1XOrig
        newMove.p1YOrig = aMove.p1YOrig
        newMove.p1XDest = aMove.p1XDest
        newMove.p1YDest = aMove.p1YDest
        newMove.p2id = aMove.p2id
        newMove.p2XOrig = aMove.p2XOrig
        newMove.p2YOrig = aMove.p2YOrig
        newMove.p2XDest = aMove.p2XDest
        newMove.p2YDest = aMove.p2YDest
        newMove.MCode = aMove.MCode
        mMoves(mCount) = newMove 'first move is 0
        mCount = mCount + 1 'mCount is 1
    End Sub

    Public Sub remove()
        'Destroy the last move
        Dim aMove As Move

        mCount = mCount - 1
        aMove = mMoves(mCount)
        aMove = Nothing
    End Sub

    Public Function last() As Move
        'return the last move
        Try
            last = mMoves(mCount - 1)
        Catch ex As Exception
            MsgBox("Can't undo")
            last = Nothing
        End Try
    End Function

    Public Sub New()
        mCount = 0
    End Sub
End Class
