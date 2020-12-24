<Serializable()>
Public Class Move
    Private mp1ID As Short 'Piece 1's move
    Private mp1XOrig As Short
    Private mp1YOrig As Short
    Private mp1XDest As Short
    Private mp1YDest As Short
    Private mp2ID As Short 'Piece 2's move e.g. captured Piece or castling
    Private mp2XOrig As Short
    Private mp2YOrig As Short
    Private mp2XDest As Short
    Private mp2YDest As Short
    Private mMCode As Char 'e.g. q for promoting, 
    '(! for check, x for take, c for castle, p for en passant) 'but DON'T OVERWRITE IMPORTANT THINGS
    'LIKE q for promoting otherwise it breaks the undoing... should have separate fields for each i guess
    'later if the Piece is on the same side maybe it's castling...
    'in en passant the destination may be empty and the pawn 1 further forward maybe the 2nd Piece
    'in promoting the  type of the Piece needs to change... but the move doesn't need to be special

    Public Property P1id() As Short
        Get
            P1id = mp1ID
        End Get
        Set(ByVal value As Short)
            mp1ID = value
        End Set
    End Property

    Public Property P1XOrig() As Short
        Get
            P1XOrig = mp1XOrig
        End Get
        Set(ByVal value As Short)
            mp1XOrig = value
        End Set
    End Property

    Public Property P1YOrig() As Short
        Get
            P1YOrig = mp1YOrig
        End Get
        Set(ByVal value As Short)
            mp1YOrig = value
        End Set
    End Property

    Public Property P1XDest() As Short
        Get
            P1XDest = mp1XDest
        End Get
        Set(ByVal value As Short)
            mp1XDest = value
        End Set
    End Property

    Public Property P1YDest() As Short
        Get
            P1YDest = mp1YDest
        End Get
        Set(ByVal value As Short)
            mp1YDest = value
        End Set
    End Property

    Public Property P2id() As Short
        Get
            P2id = mp2ID
        End Get
        Set(ByVal value As Short)
            mp2ID = value
        End Set
    End Property

    Public Property P2XOrig() As Short
        Get
            P2XOrig = mp2XOrig
        End Get
        Set(ByVal value As Short)
            mp2XOrig = value
        End Set
    End Property

    Public Property P2YOrig() As Short
        Get
            P2YOrig = mp2YOrig
        End Get
        Set(ByVal value As Short)
            mp2YOrig = value
        End Set
    End Property

    Public Property P2XDest() As Short
        Get
            P2XDest = mp2XDest
        End Get
        Set(ByVal value As Short)
            mp2XDest = value
        End Set
    End Property

    Public Property P2YDest() As Short
        Get
            P2YDest = mp2YDest
        End Get
        Set(ByVal value As Short)
            mp2YDest = value
        End Set
    End Property

    Public Property MCode() As Char
        Get
            MCode = mMCode
        End Get
        Set(ByVal value As Char)
            mMCode = value
        End Set
    End Property

    Public Sub Clear()
        mp1ID = 0
        mp1XOrig = 0
        mp1YOrig = 0
        mp1XDest = 0
        mp1YDest = 0
        mp2ID = 0
        mp2XOrig = 0
        mp2YOrig = 0
        mp2XDest = 0
        mp2YDest = 0
        mMCode = ""
    End Sub

End Class
