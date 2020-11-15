Option Strict Off
Option Explicit On
Module modChess1
    '    'The routines in this module are for general use of all Piece objects in
    '    'determining legal moves
    '    'incorporates the board 'object' to avoid having to pass it to Pieces...

    Public gBoard(MaxX, MaxY, 1) As Short 'original: 3rd dim is for id and direction
    Public gBoard2(MaxX, MaxY, 1) As Short 'copy:
    Const PIECEID As Short = 0
    Const PIECETEAM As Short = 1

    ' Function to provide unique identifiers for objects.
    Public Function GetDebugID() As Integer
        'Generate a unique Id for each object created
        '0 to 31
        Static lngDebugID As Integer
        lngDebugID += 1
        GetDebugID = lngDebugID
    End Function

    Public Sub SetGBoard(ByRef x As Byte, ByRef y As Byte, ByRef OID As Byte, ByRef d As Short)
        'the unique ID of each Piece is stored in an array to indicate its position on the board
        gBoard(x, y, PIECEID) = OID 'object identifier
        gBoard(x, y, PIECETEAM) = d 'direction
    End Sub

    Public Function IsLegal(ByRef x As Short, ByRef y As Short, ByRef d As Short) As Short
        'Check destination (defined by a pair of coordinates) - reporting existence of opponent or offboard
        'Returns 0 for empty square, same direction as moving Piece for offboard situation and direction of opponent when present.
        'Maybe should check for other illegal situations but I think that might be covered elsewhere (i.e. in scoring) i.e. move resulting in check because you're exposing the king.
        If x < 1 Or x > MaxX Or y < 1 Or y > MaxY Then 'illegal: destination is not on the board
            IsLegal = d 'tell 'em its the same direction as them so they can't move there
        Else 'tell 'em what's on the destination square normally if occupied by same you can't move there
            IsLegal = gBoard(x, y, PIECETEAM) 'may be -1, 0 or 1
        End If
    End Function
End Module