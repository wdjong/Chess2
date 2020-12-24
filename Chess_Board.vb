Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Board_NET.Board")> Public Class Board
    'A board is an 8 x 8 array number 1-8 x is across and y is number 1 at bottom to 8 on top
    'Its defined in Chess_Module1.vb as a 3 dimensional array as each location store an id (of the Piece) and a direction (of the Piece)
    'There are 2 boards for backing up the current setup before trying a move.
    Const PIECEID As Short = 0
    Const PIECETEAM As Short = 1

    Public Sub Clear()
        Dim x As Byte
        Dim y As Byte

        For x = 1 To MaxX
            For y = 1 To MaxY
                gBoard(x, y, PIECEID) = 0
                gBoard(x, y, PIECETEAM) = 0
            Next y
        Next x
    End Sub

    Public Sub CopyBoard()
        'Copy the ids and team information from the main board to backup board
        Dim x As Byte
        Dim y As Byte

        For x = 1 To MaxX 'defined in Module1
            For y = 1 To MaxY 'defined in Module1
                gBoard2(x, y, PIECEID) = gBoard(x, y, PIECEID)
                gBoard2(x, y, PIECETEAM) = gBoard(x, y, PIECETEAM)
            Next y
        Next x
    End Sub

    Public Sub RestoreBoard()
        'Copy the backup board id and team info back to the main board
        Dim x As Byte
        Dim y As Byte

        For x = 1 To MaxX
            For y = 1 To MaxY
                gBoard(x, y, PIECEID) = gBoard2(x, y, PIECEID)
                gBoard(x, y, PIECETEAM) = gBoard2(x, y, PIECETEAM)
            Next y
        Next x
    End Sub

    Public Function GetGBoardID(ByRef x As Byte, ByRef y As Byte) As Integer
        'Get the id of the Piece at this position
        GetGBoardID = gBoard(x, y, PIECEID) 'id of Piece
    End Function

    Public Function SetGBoardID(ByRef x As Byte, ByRef y As Byte, ByRef i As Byte) As Boolean
        'Set the id of the Piece at this position
        gBoard(x, y, PIECEID) = i 'id of Piece
        SetGBoardID = True
    End Function

    Public Function GetGBoardDir(ByRef x As Byte, ByRef y As Byte) As Integer
        'Get the team identification of the Piece at this position
        GetGBoardDir = gBoard(x, y, PIECETEAM) 'direction of Piece
    End Function

    Public Function SetGBoardDir(ByRef x As Byte, ByRef y As Byte, ByRef d As Short) As Boolean
        'Set the team of the Piece at this position
        gBoard(x, y, PIECETEAM) = d 'direction
        SetGBoardDir = True
    End Function

    Public Function ToString2() As String
        'Output positions to debug window
        'Not normally in use
        Dim X As Byte
        Dim Y As Byte
        Dim i As Integer
        ToString2 = ""

        For Y = 1 To MaxY
            For X = 1 To MaxX
                i = aBoard.GetGBoardID(X, (MaxY + 1) - Y) 'draw starting at top left
                If i > 0 Then 'something on it
                    If aChessPiece(i).XPos <> X Or aChessPiece(i).YPos <> (MaxY + 1) - Y Then
                        Stop 'checking that Pieces are in the same position board thinks they're in
                    End If
                End If
                If aBoard.GetGBoardDir(X, 9 - Y) = -1 Then ToString2 += "-" Else ToString2 += " "
                ToString2 += i.ToString("00")
            Next X
            ToString2 += vbCrLf
        Next Y
    End Function



End Class