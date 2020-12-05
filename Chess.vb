Option Strict Off
Option Explicit On
Imports System.IO
'Walter de Jong
'8/10/2007
'To do
'Control of central squares better than of outer
'Look for good things like forks and weight them (in measuring threat is the value of all threatened Pieces considered?)
'game stages: Perhaps use the count of moves to adjust the weightings of certain imperitives. i.e. openness... or put an upper limit on it.
'   opening: open and control
'   middle: capture and support
'   end game: homing in on king: aim to control king's square and all possible positions he can move to...
'look at the value of the exchange e.g. if I take this and they take me is that a good idea
'risk/benefit: for each Piece: If i take this, he can take that
'   subtract the values of the Pieces that threaten a Piece 
'   add the values of the Pieces that it threatens
'   add the number of the Pieces supporting a Piece
'   subtract the value of the Piece and supporting Pieces
'plex work out their next move subtract it, our next move add it etc.
'consider protection via being surrounded by pawns: sort of covered in restricting and reducing threat
'en passant
'promoting to other than queen
'computer castling
'checking human move
'more standard Algebraic notation

'look at whether threatened Pieces are supported: done: 11/8/2008; only threatened Pieces are considered

Module modChess
    Public Const MAXPIECE As Short = 32 'Total number of chess Pieces in play
    Public Const BOARDTOP As Short = 0
    Public Const BOARDLEFT As Short = 0
    Public Const MaxX As Short = 8
    Public Const MaxY As Short = 8
    Public PWIDTH As Integer = 32 'pixels
    Public PHEIGHT As Integer '= 32 'pixels: width of a board picture thing for mouse calculations
    Public PIECEWIDTH As Short '= 480 'PWIDTH * 15 '32 * 15 = 480 'twips: both Piece and square
    Public PIECEHEIGHT As Short '= 480 'PHEIGHT * 15 '
    Public BOARDBOTTOM As Integer '= (MaxY + 1) * PIECEHEIGHT
    Public BBOTTOM As Integer '= '(MaxY + 1) * PHEIGHT
    Public aBoard As Board
    Public aChessPiece(MAXPIECE) As ChessPiece 'NB array index must match control index on the form for the corresponding Piece
    Public PieceArray(MAXPIECE) As System.Windows.Forms.PictureBox
    Public aGame As New Game
    Public aWhite As New Player(1)
    Public Turn As Short 'Who's turn is it? Black/Blue = -1 Top going down or White/Red = 1 Bottom going up
    Public Const TOPGOINGDOWN As Short = -1 'Direction/ Turn constant
    Public Const BOTTOMGOINGUP As Short = 1
    Public PlayerisHuman(2) As Boolean 'bit ugly but make player 0 top and player 2 bottom
    Const captureMove As Byte = 2 ' e.g. aChessPiece(aPiece).getscore(aMove) = captureMove
    Public Thoughts As String
    Public Property displayMoves As Boolean = True

    Sub DebugBoardPrint()
        'Output positions to debug window
        'Not normally in use
        Dim X As Byte
        Dim Y As Byte
        Dim i As Integer

        For Y = 1 To MaxY
            For X = 1 To MaxX
                i = aBoard.GetGBoardID(X, (MaxY + 1) - Y) 'draw starting at top left
                If i > 0 Then 'something on it
                    If aChessPiece(i).XPos <> X Or aChessPiece(i).YPos <> (MaxY + 1) - Y Then
                        Stop 'checking that Pieces are in the same position board thinks they're in
                    End If
                End If
                If aBoard.GetGBoardDir(X, 9 - Y) = -1 Then System.Diagnostics.Debug.Write("-") Else System.Diagnostics.Debug.Write(" ")
                System.Diagnostics.Debug.Write(i.ToString("00"))
            Next X
            System.Diagnostics.Debug.WriteLine("")
        Next Y
    End Sub

    Sub DoBestMove(ByVal pbest As Byte, ByVal mbest As Byte)
        'Create move object and use it to update the board, the Pieces, record the move and display the Piece
        Dim aMove As New Move 'create a new move object for storing the current move in 
        Dim algNotDest As String

        aChessPiece(pbest).CheckMoves() 'Updates MaxMove of that Piece (necessary to restore correct move?!)
        aMove.P1id = pbest 'record the first Piece's move
        aMove.P1XOrig = aChessPiece(pbest).XPos
        aMove.P1YOrig = aChessPiece(pbest).YPos
        aMove.P1XDest = aChessPiece(pbest).GetMoveX(mbest)
        aMove.P1YDest = aChessPiece(pbest).GetMoveY(mbest)
        aMove.P2id = aBoard.GetGBoardID(aMove.P1XDest, aMove.P1YDest) 'Check what's on board at piece destination
        If aMove.P2id > 0 Then 'there's a Piece being taken
            aMove.P2XOrig = aMove.P1XDest 'register the origin of the piece being taken in the move object
            aMove.P2YOrig = aMove.P1YDest
            Take(aMove.P1XDest, aMove.P1YDest) 'update the board and taken Piece if necessary
        End If
        aChessPiece(aMove.P1id).Move(aMove.P1XDest, aMove.P1YDest) 'update the moving Piece's location (and the board)
        If Ispromoting(aMove.P1id, aMove.P1YDest) Then 'it realises it's reached the 8th rank
            Queening(aMove.P1id) 'change into a queen
            aMove.MCode = "q" 'in case we need to undo
        End If
        aGame.Add(aMove) 'add the move to the array
        If displayMoves = True Then
            algNotDest = GetAlgebraicNotation(aMove)
            DoHistory(algNotDest, aChessPiece(aMove.P1id).Direction)
            My.Forms.FrmBoard.ShowPiece(aMove.P1id) 'ensure the Piece is visible and in the right location on the board
            My.Forms.FrmBoard.StatusStrip1.Items("statFrom").Text = "From: " & aChessPiece(aMove.P1id).AlgNot & Chr(96 + aMove.P1XOrig) & aMove.P1YOrig
            My.Forms.FrmBoard.StatusStrip1.Items("statTo").Text = "To: " & algNotDest
        End If
    End Sub

    Friend Function GetAlgebraicNotation(aMove As Move) As String
        'Need to happen before Turn = -Turn
        'https://www.chesscentral.com/pages/learn-chess-play-better/chess-notation-keeping-score.html#:~:text=%3A%20Kxe2%2C%20Qxh5.-,8.,For%20example%3A%20dxe5%20or%20hxg6.
        'Only thing not happening now is if there's two knights or two rooks can move to same square (so it's ambiguous) could check all possible moves for other piece
        Dim notation As String

        If aChessPiece(aMove.P1id).AlgNot <> "P" Then
            notation = aChessPiece(aMove.P1id).AlgNot 'piece name for 
        Else
            notation = Chr(96 + aMove.P1XOrig) 'file for pawn
        End If
        If aMove.P2id > 0 Then 'a piece is being taken
            notation += "#" 'capture
        End If
        notation += Chr(96 + aMove.P1XDest) & aMove.P1YDest 'destination square
        If IsInCheck(-Turn) Then 'warn the opponent that they're in check 'just if human I guess
            notation += "+" 'signifies Check
            If PlayerisHuman(1 + -Turn) Then
                MsgBox("Check")
            End If
        End If
        If IsCheckMate(-Turn) Then 'warn the opponent that they're in check 'just if human I guess
            notation += "#" 'signifies Check
            If PlayerisHuman(1 + -Turn) Then
                MsgBox("Checkmate")
            End If
        End If
        GetAlgebraicNotation = notation
    End Function

    Sub DoCastle(ByVal aMove As Move)
        'carry out the castling
        aMove.P2XOrig = 5 'king's always on 5
        aMove.P2YOrig = aMove.P1YOrig 'either 1 or 8
        aMove.P1YDest = aMove.P1YOrig
        aMove.P2YDest = aMove.P1YOrig
        If aMove.P1XOrig = 1 Then 'queenside
            aMove.P1XDest = 4
            aMove.P2XDest = 3
        Else 'kingside
            aMove.P1XDest = 6
            aMove.P2XDest = 7
        End If
        aChessPiece(aMove.P1id).Move(aMove.P1XDest, aMove.P1YDest) 'Move rook
        aChessPiece(aMove.P2id).Move(aMove.P2XDest, aMove.P2YDest) 'Move king
        If IsInCheck(Turn) Then 'Check again for legality after the move
            MsgBox("You're in check. Either you moved into it or you didn't move out of it.")
            aChessPiece(aMove.P1id).Move(aMove.P1XOrig, aMove.P1YOrig) 'move Piece back
            aChessPiece(aMove.P2id).Move(aMove.P2XOrig, aMove.P2YOrig) 'move taken Piece back
        Else
            aGame.Add(aMove) 'add the move to the array
            Turn = -Turn
        End If
    End Sub

    Sub CheckOpenness(ByRef aPiece As ChessPiece, ByVal currentPlayersDirection As Short, ByRef ourOpenness As Short, ByRef theirOpenness As Short)
        'Calculates the openness which (apart from pawns) equates to controlled squares
        Dim maxLegal As Short

        Try
            maxLegal = aPiece.CheckMoves 'get number of legal moves
            If aPiece.Direction = currentPlayersDirection Then
                ourOpenness += maxLegal  'this is also roughly the same as contolled spaces
            Else
                theirOpenness += maxLegal
            End If
        Catch ex As Exception
            Stop
            MsgBox("CheckOpenness: " & ex.Message)
        End Try
    End Sub

    Sub CheckSupport(ByRef aPiece As ChessPiece, ByVal currentPlayersDirection As Short, ByRef ourSuppCount As Integer, ByRef theirSuppCount As Integer)
        'Find the value of all the supported Pieces
        'If we use Threatened property that's set in CheckThreats so that would have to have been done first.
        Const takePieceValPwr As Byte = 1 'this adds the capture Piece value to this power
        Dim maxSupport As Byte
        Dim aMove As Byte 'each move
        Dim aDestX As Byte 'destination x
        Dim aDestY As Byte 'destination y
        Dim aSupportedID As Byte 'id of threatened Piece
        Dim aValue As Byte

        Try
            'If My.Settings.ShareThoughts and aPiece.Watch Then
            '            Cout("CheckSupport: " & ourSuppCount & " " & theirSuppCount & vbNewLine)
            'End If
            maxSupport = aPiece.CheckSupport 'get an array of supporting moves i.e. moves where this could take the piece that's replaced one of it's own
            For aMove = 1 To maxSupport 'each supporting possibility
                aDestX = aPiece.GetSuppX(aMove) 'determine the co-ordinates of the supported Piece
                aDestY = aPiece.GetSuppY(aMove)
                aSupportedID = aBoard.GetGBoardID(aDestX, aDestY) 'get the id of the supported Piece
                aValue = aChessPiece(aSupportedID).Value * takePieceValPwr 'get value of supported Piece
                aChessPiece(aSupportedID).Supported += 1 'used in checkUnsupported routine
                'If aChessPiece(aSupportedID).Threatened = 0 Then aValue = 0 'Note: Assumes threats have been assessed and doesn't care about supporting unthreatened pieces...
                If aChessPiece(aSupportedID).AlgNot = "K" Then aValue = 0 'no point supporting the king
                If aPiece.Direction = currentPlayersDirection Then 'it's our Piece
                    ourSuppCount += aValue * takePieceValPwr 'Find the total level of support for other Pieces
                Else 'it's their (the opponent's) Piece
                    theirSuppCount += aValue * takePieceValPwr 'beTakenCount = beTakenCount + aValue * AvoidingThreat 'add the value of the Piece it can take to the opponents take count
                End If
            Next aMove
            'If My.Settings.ShareThoughts and aPiece.Watch Then
            '            Cout("CheckSupport: " & ourSuppCount & " " & theirSuppCount & vbNewLine)
            'End If
        Catch ex As Exception
            MsgBox("CheckSupport: " & ex.Message)
        End Try
    End Sub

    Sub CheckThreats(ByRef aPiece As ChessPiece, ByVal currentPlayersDirection As Short, ByRef takeCount As Integer, ByRef betakenCount As Integer)
        'This is called for every piece to get the values of what is being threatened.
        'Calculates the total value of all threatened Pieces for each player
        Const takePieceValPwr As Byte = 1 'this adds the capture Piece value to this power
        Const advantageousExchangeWt As Byte = 1 '
        Dim aMove As Byte 'each move
        Dim aDestX As Byte 'destination x
        Dim aDestY As Byte 'destination y
        Dim aThreatenedID As Byte 'id of threatened Piece
        Dim aValue As Byte

        Try
            'If My.Settings.ShareThoughts aPiece.Watch Then
            '     Cout("CheckThreats: " & takeCount & " " & betakenCount & vbNewLine)
            'End If
            For aMove = 1 To aPiece.CheckMoves 'maxLegal 'each legal move
                If aPiece.GetScore(aMove) = captureMove Then 'find the value of threatened Piece
                    aDestX = aPiece.GetMoveX(aMove) 'determine the co-ordinates of the destination
                    aDestY = aPiece.GetMoveY(aMove)
                    aThreatenedID = aBoard.GetGBoardID(aDestX, aDestY) 'get the id of the destination Piece
                    aValue = aChessPiece(aThreatenedID).Value * takePieceValPwr 'weight the value of threatened Piece
                    aChessPiece(aThreatenedID).Threatened += 1 'keep track of how many attackers there are (NOT CURRENTLY USED?)
                    If aPiece.Direction = currentPlayersDirection Then 'this Piece belongs to the player whose turn it is
                        If aPiece.Value < aChessPiece(aThreatenedID).Value Then 'weight advantageous exchanges
                            takeCount += aValue  'add the value of the Piece it can take to takeCount
                        Else
                            takeCount += aValue * advantageousExchangeWt
                        End If
                    Else 'the opponent's Piece
                        betakenCount += aValue 'add the value of the Piece it can take to the opponents take count
                    End If
                End If
            Next aMove
            'If My.Settings.ShareThoughts aPiece.Watch Then
            '     Cout("CheckThreats: " & takeCount & " " & betakenCount & vbNewLine)
            'End If
        Catch ex As Exception
            MsgBox("checkThreats: " & ex.Message)
        End Try
    End Sub

    Sub CheckUnsupported(ByVal currentPlayersDirection As Short, ByRef ourUnSuppValue As Integer, ByRef theirUnSuppValue As Integer)
        'works out the total value & level of unsupported Pieces
        'probably should use threatened unsupported Pieces
        'The loop relating to support needs to have completed for all Pieces
        'REQUIRES: CheckSupport to have been run
        Const takePieceValPwr As Byte = 1.5 'this adds the capture Piece value to this power
        Dim aPiece As Byte

        Try
            For aPiece = 1 To MAXPIECE
                'If My.Settings.ShareThoughts and aChessPiece(aPiece).Watch Then
                '        Cout(vbNewLine & "CheckUnsupported c:" & ourUnSuppValue & " h:" & theirUnSuppValue & vbNewLine)
                'End If
                If aChessPiece(aPiece).OnBoard Then
                    If aChessPiece(aPiece).Direction = currentPlayersDirection Then 'just look at our Pieces
                        If aChessPiece(aPiece).Supported = 0 Then 'ignore kings as they don't need supporting
                            If aChessPiece(aPiece).AlgNot <> "K" Then
                                ourUnSuppValue += aChessPiece(aPiece).Value * takePieceValPwr
                            End If
                        End If
                    Else 'their Pieces
                        If aChessPiece(aPiece).AlgNot <> "K" Then 'ignore kings as they don't need supporting
                            If aChessPiece(aPiece).Supported = 0 Then
                                theirUnSuppValue += aChessPiece(aPiece).Value * takePieceValPwr
                            End If
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox("CheckUnsupported: " & ex.Message)
        End Try

    End Sub

    Sub CheckTotalValue(ByRef aPiece As ChessPiece, ByVal currentPlayersDirection As Short, ByRef ourPieceTtlValue As Integer, ByRef theirPieceTtlValue As Integer)
        'Find the total value of all remaining Pieces
        Const takePieceValPwr As Byte = 2

        If aPiece.Direction = currentPlayersDirection Then 'just look at our Pieces
            ourPieceTtlValue += aPiece.Value * takePieceValPwr
        Else 'their Pieces
            theirPieceTtlValue += aPiece.Value * takePieceValPwr
        End If
    End Sub

    Sub DoHistory(ByVal aDest As String, ByVal aDir As Short)
        'update history display on console out
        'aDest is in the form of Algnot + alpha file + rank e.g. KE5 is Knight to E5
        'aDir is a direction e.g. 1 or -1
        Cout(aDest)
        If aDir = -1 Then
            Cout(vbCrLf)
        Else
            Cout("        ")
        End If
    End Sub

    Sub GetBestMove(ByRef pBest As Byte, ByRef mBest As Byte)
        'Basic Move weightings are used in conjunction with Board/Position strength assessments.
        Const checkIncentive As Byte = 1
        Const takeIncentive As Byte = 2 'this adds the capture Piece value to this power i.e. Value^2
        Const promoteIncentive As Byte = 10
        Const deterRepetition As Integer = 10
        Const checkMateIncentive As Byte = 20

        Dim p As Byte 'each Piece
        Dim pf As Byte 'first Piece
        Dim pl As Byte 'last Piece
        Dim ml As Byte 'number of valid moves
        Dim m As Byte '1 to number of moves
        Dim s As Long 'score for each move
        Dim sl As Long 'maximum score
        Dim result As Byte
        Dim inCheck As Boolean
        Dim repetitions As Integer

        sl = Integer.MinValue ' -10000 'best score so far
        pBest = 0 'best Piece index
        mBest = 0 'best move index
        If Turn = TOPGOINGDOWN Then
            pf = 17
            pl = MAXPIECE
        Else 'BOTTOMGOINGUP
            pf = 1
            pl = 16
        End If
        For p = pf To pl 'For each Piece in team
            If My.Settings.ShareThoughts And aChessPiece(p).Watch Then
                Thoughts = ""
            End If
            If aChessPiece(p).OnBoard Then 'when a Piece is taken OnBoard is set to false
                ml = aChessPiece(p).CheckMoves 'count legal moves a Piece can make and add them to an array of moves in the Piece object
                If My.Settings.ShareThoughts And aChessPiece(p).Watch Then
                    Thoughts += "CheckMoves: " & ml & vbNewLine
                End If

                If ml > 0 Then
                    For m = 1 To ml 'For each possible move 'The list of legal moves are stored in the Piece
                        aBoard.CopyBoard() 'each board object can remember a copy of itself
                        'Try Piece p's mth move; which may be capturing something
                        TakeNoShow(aChessPiece(p).GetMoveX(m), aChessPiece(p).GetMoveY(m)) 'perform capture
                        aChessPiece(p).Move(m)
                        s = Score(Turn) 'Derive a score based on the resulting board arrangement
                        If IsInCheck(-Turn) Then s += checkIncentive 'Add some incentive to check opponent
                        If IsCheckMate(-Turn) Then s += checkMateIncentive 'Add some incentive to check opponent
                        inCheck = IsInCheck(Turn) 'Ensure you move out of check
                        If inCheck Then
                            Debug.Print("puts in check: " & aChessPiece(p).AlgNot & Chr(64 + aChessPiece(p).GetMoveX(m)) & "" & aChessPiece(p).GetMoveY(m) & vbNewLine)
                        End If
                        aBoard.RestoreBoard()
                        PiecesFromBoard() 'recreate Piece positions from board
                        result = aChessPiece(p).CheckMoves 'ensure this Piece has current moves again
                        If aChessPiece(p).GetScore(m) = captureMove Then 'Add some incentive to take the opponent
                            s += aChessPiece(aBoard.GetGBoardID(aChessPiece(p).GetMoveX(m), aChessPiece(p).GetMoveY(m))).Value ^ takeIncentive 'get the value of the opponents Piece
                            If My.Settings.ShareThoughts And aChessPiece(p).Watch Then
                                Thoughts += "Capture move: " & aChessPiece(aBoard.GetGBoardID(aChessPiece(p).GetMoveX(m), aChessPiece(p).GetMoveY(m))).Value ^ takeIncentive & vbNewLine
                            End If
                        End If
                        If Ispromoting(p, aChessPiece(p).GetMoveY(m)) Then
                            s += promoteIncentive 'Add some incentive to queen
                            If My.Settings.ShareThoughts And aChessPiece(p).Watch Then
                                Thoughts += "Promote incentive added: " & vbNewLine
                            End If
                        End If
                        repetitions = CountRepetition(p, m)
                        s -= deterRepetition * repetitions
                        If s > sl And Not inCheck And repetitions < 3 Then
                            sl = s 'Keep track of the best score
                            pBest = p 'the best Piece
                            mBest = m 'the best move
                            If My.Settings.ShareThoughts And aChessPiece(p).Watch Then
                                Thoughts += vbNewLine & "Piece: " & aChessPiece(p).AlgNot & Chr(64 + aChessPiece(p).GetMoveX(m)) & "" & aChessPiece(p).GetMoveY(m) & " Score: " & s & "--> "
                            End If
                        End If
                    Next m
                End If
                If My.Settings.ShareThoughts And aChessPiece(p).Watch Then
                    Cout(Thoughts)
                End If
            End If
        Next p
    End Sub

    Function CountRepetition(p As Byte, m As Byte) As Integer
        CountRepetition = 0

        If aGame.Count > 0 Then
            For i As Integer = 0 To aGame.Count - 1
                If aGame.Moves(i).P1XOrig = aChessPiece(p).XPos _
            And aGame.Moves(i).P1YOrig = aChessPiece(p).YPos _
            And aGame.Moves(i).P1XDest = aChessPiece(p).GetMoveX(m) _
            And aGame.Moves(i).P1YDest = aChessPiece(p).GetMoveY(m) _
            And aGame.Moves(i).P1id = p Then
                    CountRepetition += 1
                End If
            Next
        End If
    End Function

    Function IsInCheck(ByVal aDir As Short) As Boolean
        'aDir is the colour of the king that could be in check i.e. 1 BottomGoingUp Red/White or -1 TopGoingDown Blue/Black
        'Check if any moves for opposition Pieces are threatening the King
        Dim p As Byte 'each Piece
        Dim pf As Byte 'first Piece
        Dim pl As Byte 'last Piece
        Dim ml As Byte 'number of valid moves
        Dim m As Byte '1 to number of moves
        Dim aX As Byte
        Dim aY As Byte
        Dim aID As Byte

        IsInCheck = False
        Try
            If aDir = TOPGOINGDOWN Then 'blue: identify the range of opposition Pieces
                pf = 1 'look at red Pieces
                pl = 16
            Else 'BOTTOMGOINGUP
                pf = 17
                pl = MAXPIECE
            End If
            For p = pf To pl 'For each Piece in the opponent's team
                ml = aChessPiece(p).CheckMoves 'find how many moves the opponent's Piece can make
                If ml > 0 Then
                    For m = 1 To ml 'For each possible move
                        If aChessPiece(p).GetScore(m) > 1 Then 'something can be taken
                            aX = aChessPiece(p).GetMoveX(m) 'determine the co-ordinates of the destination
                            aY = aChessPiece(p).GetMoveY(m)
                            aID = aBoard.GetGBoardID(aX, aY) 'get the id of the destination Piece
                            If aChessPiece(aID).AlgNot() = "K" Then 'it's the king
                                IsInCheck = True
                                If My.Settings.ShareThoughts Then
                                    Thoughts += "Check: " & aChessPiece(p).AlgNot & " is threatening the king."
                                End If
                            End If
                        End If
                    Next m
                End If
            Next p
        Catch ex As Exception
            MsgBox("isInCheck: " & ex.Message)
        End Try
    End Function

    Function IsCheckMate(ByVal aDir As Short) As Boolean
        Dim ml As Byte 'number of valid moves
        Dim p As Byte

        IsCheckMate = False
        If IsInCheck(aDir) Then
            If aDir = BOTTOMGOINGUP Then p = 5 Else p = 29
            ml = aChessPiece(p).CheckMoves 'find how many moves the opponent's Piece can make
            If ml = 0 Then
                IsCheckMate = True
            End If
        End If
    End Function

    Function IsLegal1(ByVal aMove As Move) As Boolean
        'Initial check on the validity of the move
        IsLegal1 = True
        If Not aGame.Playing Then
            IsLegal1 = False 'we're not in a game
        ElseIf Turn <> aChessPiece(aMove.P1id).Direction Then
            IsLegal1 = False 'moving wrong team: not your go
        ElseIf aMove.P1XDest > MaxX Or aMove.P1YDest > MaxY Or aMove.P1XDest < 1 Or aMove.P1YDest < 1 Then
            IsLegal1 = False 'moving off board
        ElseIf aMove.P2id > 0 Then
            If aChessPiece(aMove.P1id).Direction = aChessPiece(aMove.P2id).Direction Then
                IsLegal1 = False 'trying to take own Piece
            End If
        End If
    End Function

    Function Ispromoting(ByVal aPiece As Byte, ByVal aDestY As Byte) As Boolean
        'since a pawn can't retreat and always starts on the second rank it's okay to test for them being on 1 or 8
        Ispromoting = False
        If aChessPiece(aPiece).AlgNot = "P" Then
            If aDestY = 1 Or aDestY = 8 Then 'promoting
                Ispromoting = True
            End If
        End If
    End Function

    Sub PiecesFromBoard()
        'Having restored the board
        'update each Piece based on it's position on the board
        Dim X As Byte
        Dim Y As Byte
        Dim i As Integer

        For i = 1 To 32
            aChessPiece(i).XPos = 0
            aChessPiece(i).YPos = 0
            aChessPiece(i).OnBoard = False
            aChessPiece(i).Direction = 0
            aChessPiece(i).Supported = 0 'set in scoring routine
            aChessPiece(i).Threatened = 0 'set in scoring routine
        Next
        For X = 1 To MaxX 'each board position
            For Y = 1 To MaxY
                i = aBoard.GetGBoardID(X, Y) 'NB Piece IDs are 1 to 32
                If i > 0 Then 'there is a Piece is on it (Piece i)
                    aChessPiece(i).XPos = X 'set the position of the Piece
                    aChessPiece(i).YPos = Y 'set location of Piece
                    aChessPiece(i).OnBoard = True
                    If i > 16 Then
                        aChessPiece(i).Direction = TOPGOINGDOWN
                    Else
                        aChessPiece(i).Direction = BOTTOMGOINGUP
                    End If
                End If
            Next Y
        Next X
    End Sub

    Sub Queening(ByVal p As Byte)
        'Change Piece p into a queen
        'Note leave out the moving and showing as that's handled elsewhere
        Dim aQueen As New Queen With {
            .Direction = aChessPiece(p).Direction,
            .ObjIndex = p
        } 'sets AlgNot and value
        aQueen.Move(aChessPiece(p).XPos, aChessPiece(p).YPos)
        aChessPiece(p) = aQueen 'change into a queen
        If p > 16 Then
            PieceArray(p).Image = My.Resources.queen1 'PieceArray(28).Image
        Else
            PieceArray(p).Image = My.Resources.queen 'PieceArray(4).Image
        End If
    End Sub

    Private Function Score(ByVal d As Short) As Long
        'Check the relative strength of the two sides
        'given the direction d is the direction of the current player
        'Returns integer.minvalue if there would be no move
        'Because very large numbers can be generated use the following value with care... They are basically pairs of characteristics from our point of view and theirs.
        Const Threatening As Byte = 0.9 'aggressiveness - increase our threatening
        Const AvoidingThreat As Byte = 2 'wimpishness - reduce their threatening
        Const Supporting As Byte = 1 'supporting our own Pieces - increase our support
        Const Undermine As Byte = 1 'reduce their support
        Const Control As Byte = 1 'aim to control more of the board
        Const Frustrate As Byte = 1 'aim to reduce their mobility and control
        Const Keep As Byte = 1 'aim to keep as many Pieces as possible
        Const Erode As Byte = 1 'aim to reduce their Pieces
        Const Consolidate As Byte = 8 'aim to avoid leaving Piece unsupported
        Const Isolate As Byte = 1 'aim to isolate their Pieces leaving them unsupported
        Dim aPiece As Byte 'for each Piece
        Dim takeCount As Integer = 0 'sum the value of all their threatened Pieces from proposed position
        Dim beTakenCount As Integer = 0 'sum of the value of our threatened Pieces from proposed position
        Dim ourSuppCount As Integer = 0 'sum the value of all our supported Pieces from proposed position
        Dim theirSuppCount As Integer = 0
        Dim ourUnSuppValue As Integer = 0 'sum the value of all the unsupported Pieces from proposed position
        Dim theirUnSuppValue As Integer = 0 'sum the value of all their unsupported Pieces
        Dim ourPieceTtlValue As Integer = 0 'on board
        Dim theirPieceTtlValue As Integer = 0 'on board
        Dim ourOpenness As Short = 0 'number of possible moves (roughly the same as controlled spaces)
        Dim theirOpenness As Short = 0 'for the opponent

        Score = 0
        Try
            For aPiece = 1 To MAXPIECE 'for each Piece (in it's position after a proposed move)
                If aChessPiece(aPiece).OnBoard() Then
                    CheckTotalValue(aChessPiece(aPiece), d, ourPieceTtlValue, theirPieceTtlValue)
                    CheckOpenness(aChessPiece(aPiece), d, ourOpenness, theirOpenness) 'find who's controlling more
                    CheckThreats(aChessPiece(aPiece), d, takeCount, beTakenCount) 'find who's threatening who
                    CheckSupport(aChessPiece(aPiece), d, ourSuppCount, theirSuppCount) 'find out what's supported
                End If
            Next aPiece
            CheckUnsupported(d, ourUnSuppValue, theirUnSuppValue) 'uses .supported property set in checkSupport (therefore can't be in same loop)
            If ourOpenness = 0 Then
                Score = Long.MinValue
            Else
                'Note: there's also some scoring above as we add the value of the Pieces to the takeCount
                'and back in the routine that called this some score relating to the specific move
                Score = takeCount * Threatening - beTakenCount * AvoidingThreat
                Score = Score + ourPieceTtlValue * Keep - theirPieceTtlValue * Erode
                Score = Score + ourSuppCount * Supporting - theirSuppCount * Undermine
                Score = Score + ourOpenness * Control - theirOpenness * Frustrate
                Score = Score - ourUnSuppValue * Consolidate + theirUnSuppValue * Isolate
            End If
            If My.Settings.ShareThoughts Then
                Thoughts += "TakeCnt +" & takeCount * Threatening & " -" & beTakenCount * AvoidingThreat & vbNewLine
                Thoughts += "TtlVal +" & ourPieceTtlValue * Keep & " -" & theirPieceTtlValue * Erode & vbNewLine
                Thoughts += "Support +" & ourSuppCount * Supporting & " -" & theirSuppCount * Undermine & vbNewLine
                Thoughts += "Openness +" & ourOpenness * Control & " -" & theirOpenness * Frustrate & vbNewLine
                Thoughts += "UnSupp -" & ourUnSuppValue * Consolidate & " +" & theirUnSuppValue * Isolate & vbNewLine
            End If
        Catch ex As Exception
            MsgBox("Score: " & ex.Message)
        End Try
    End Function

    Sub SetupPieces()
        'Create piece objects, set direction and index, set position and assign control/image
        Dim p As Short

        For p = 1 To MAXPIECE
            Select Case p
                Case 1 'Create object, set direction and index
                    aChessPiece(p) = New Rook With {
                        .Direction = BOTTOMGOINGUP,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move(1, 1) 'Set position
                    PieceArray(1) = FrmBoard.Piece_1 'assign control (image)
                Case 2
                    aChessPiece(p) = New Knight With {
                        .Direction = BOTTOMGOINGUP,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move(2, 1)
                    PieceArray(2) = FrmBoard.Piece_2
                Case 3
                    aChessPiece(p) = New Bishop With {
                        .Direction = BOTTOMGOINGUP,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move(3, 1)
                    PieceArray(3) = FrmBoard.Piece_3
                Case 4
                    aChessPiece(p) = New Queen With {
                        .Direction = BOTTOMGOINGUP,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move(4, 1)
                    PieceArray(4) = FrmBoard.Piece_4
                Case 5
                    aChessPiece(p) = New King With {
                        .Direction = BOTTOMGOINGUP,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move(5, 1)
                    PieceArray(5) = FrmBoard.Piece_5
                Case 6
                    aChessPiece(p) = New Bishop With {
                        .Direction = BOTTOMGOINGUP,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move(6, 1)
                    PieceArray(6) = FrmBoard.Piece_6
                Case 7
                    aChessPiece(p) = New Knight With {
                        .Direction = BOTTOMGOINGUP,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move(7, 1)
                    PieceArray(7) = FrmBoard.Piece_7
                Case 8
                    aChessPiece(p) = New Rook With {
                        .Direction = BOTTOMGOINGUP,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move(8, 1)
                    PieceArray(8) = FrmBoard.Piece_8
                Case 9 To 16
                    aChessPiece(p) = New Pawn With {
                        .Direction = BOTTOMGOINGUP,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move((p - 1) Mod 8 + 1, 2)
                    PieceArray(9) = FrmBoard.Piece_9
                    PieceArray(10) = FrmBoard.Piece_10
                    PieceArray(11) = FrmBoard.Piece_11
                    PieceArray(12) = FrmBoard.Piece_12
                    PieceArray(13) = FrmBoard.Piece_13
                    PieceArray(14) = FrmBoard.Piece_14
                    PieceArray(15) = FrmBoard.Piece_15
                    PieceArray(16) = FrmBoard.Piece_16
                    PieceArray(p).Image = My.Resources.pawn 'PieceArray(9).Image
                Case 17 To 24
                    aChessPiece(p) = New Pawn With {
                        .Direction = TOPGOINGDOWN,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move((p - 1) Mod 8 + 1, 7)
                    PieceArray(17) = FrmBoard.Piece_17
                    PieceArray(18) = FrmBoard.Piece_18
                    PieceArray(19) = FrmBoard.Piece_19
                    PieceArray(20) = FrmBoard.Piece_20
                    PieceArray(21) = FrmBoard.Piece_21
                    PieceArray(22) = FrmBoard.Piece_22
                    PieceArray(23) = FrmBoard.Piece_23
                    PieceArray(24) = FrmBoard.Piece_24
                    PieceArray(p).Image = My.Resources.pawn1 'as they might have been queened
                Case 25
                    aChessPiece(p) = New Rook With {
                        .Direction = TOPGOINGDOWN,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move(1, 8)
                    PieceArray(25) = FrmBoard.Piece_25
                Case 26
                    aChessPiece(p) = New Knight With {
                        .Direction = TOPGOINGDOWN,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move(2, 8)
                    PieceArray(26) = FrmBoard.Piece_26
                Case 27
                    aChessPiece(p) = New Bishop With {
                        .Direction = TOPGOINGDOWN,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move(3, 8)
                    PieceArray(27) = FrmBoard.Piece_27
                Case 28
                    aChessPiece(p) = New Queen With {
                        .Direction = TOPGOINGDOWN,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move(4, 8)
                    PieceArray(28) = FrmBoard.Piece_28
                    aChessPiece(p).Watch = True
                Case 29
                    aChessPiece(p) = New King With {
                        .Direction = TOPGOINGDOWN,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move(5, 8)
                    PieceArray(29) = FrmBoard.Piece_29
                Case 30
                    aChessPiece(p) = New Bishop With {
                        .Direction = TOPGOINGDOWN,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move(6, 8)
                    PieceArray(30) = FrmBoard.Piece_30
                Case 31
                    aChessPiece(p) = New Knight With {
                        .Direction = TOPGOINGDOWN,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move(7, 8)
                    PieceArray(31) = FrmBoard.Piece_31
                Case 32
                    aChessPiece(p) = New Rook With {
                        .Direction = TOPGOINGDOWN,
                        .ObjIndex = p
                    }
                    aChessPiece(p).Move(8, 8)
                    PieceArray(32) = FrmBoard.Piece_32
            End Select
            'If displayMoves = True Then
            '    My.Forms.FrmBoard.ShowPiece(p)
            'End If
        Next p
    End Sub

    Sub ShowPieces()
        For p As Integer = 1 To 32
            My.Forms.FrmBoard.ShowPiece(p)
        Next p
        Application.DoEvents()
    End Sub

    Function Take(ByRef X As Byte, ByRef Y As Byte) As Boolean
        'Call this prior to moving.
        Dim p As Byte

        Take = False
        On Error GoTo errTake
        p = aBoard.GetGBoardID(X, Y)
        If p > 0 Then
            aChessPiece(p).Remove()
            If displayMoves = True Then
                PieceArray(p).Visible = False 'make invisible
                Beep()
            End If
            Take = True 'affirm that take has occurred
        End If
exitTake:
        Exit Function
errTake:
        MsgBox("Take: " & Err.Description)
        Resume exitTake
    End Function

    Sub TakeNoShow(ByRef X As Byte, ByRef Y As Byte)
        'Call this prior to moving (called from cmdMove_Click).
        Dim p As Byte

        On Error GoTo errTake
        p = aBoard.GetGBoardID(X, Y)
        If p > 0 Then
            aChessPiece(p).Remove()
        End If
exitTake:
        Exit Sub
errTake:
        MsgBox("TakeNoShow: " & Err.Description)
        Resume exitTake
    End Sub

    Sub UnQueen(ByVal p As Byte)
        'Change Piece p into a pawn
        'For use when undoing a promoting
        Dim aPawn As New Pawn With {
            .Direction = aChessPiece(p).Direction,
            .ObjIndex = p
        } 'sets AlgNot and value
        aPawn.Move(aChessPiece(p).XPos, aChessPiece(p).YPos)
        aChessPiece(p) = aPawn 'change into a pawn
        If p > 16 Then 'blue
            PieceArray(p).Image = My.Resources.pawn1 'PieceArray(17).Image
        Else 'red
            PieceArray(p).Image = My.Resources.pawn 'PieceArray(9).Image
        End If
    End Sub

    Sub SaveResult(Score As Double)
        Dim strFile As String = "results.txt"
        Dim resultLine As String = DateTime.Now
        Dim fileExists As Boolean = File.Exists(strFile)

        resultLine += ", " & aWhite.CheckIncentive
        resultLine += ", " & aWhite.TakeIncentive 'this adds the capture Piece value to this power i.e. Value^2
        resultLine += ", " & aWhite.PromoteIncentive
        resultLine += ", " & aWhite.DeterRepetition
        resultLine += ", " & aWhite.CaptureMove ' e.g. aChessPiece(aPiece).getscore(aMove) = captureMove
        resultLine += ", " & aWhite.Threatening 'aggressiveness - increase our threatening
        resultLine += ", " & aWhite.AvoidingThreat 'wimpishness - reduce their threatening
        resultLine += ", " & aWhite.Supporting 'supporting our own Pieces - increase our support
        resultLine += ", " & aWhite.Undermine
        resultLine += ", " & aWhite.Control 'aim to control more of the board
        resultLine += ", " & aWhite.Frustrate 'aim to reduce their mobility and control
        resultLine += ", " & aWhite.Keep 'aim to keep as many Pieces as possible
        resultLine += ", " & aWhite.Erode 'aim to reduce their Pieces
        resultLine += ", " & aWhite.Consolidate 'aim to avoid leaving Piece unsupported
        resultLine += ", " & aWhite.Isolate 'aim to isolate their Pieces leaving them unsupported
        resultLine += ", " & aWhite.CheckMateIncentive 'Achieve for CheckMate
        resultLine += ", " & Score

        Using sw As New StreamWriter(File.Open(strFile, FileMode.OpenOrCreate))
            sw.WriteLine(resultLine)
        End Using
    End Sub
End Module