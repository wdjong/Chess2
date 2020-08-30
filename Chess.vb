Option Strict Off
Option Explicit On
'Walter de Jong
'8/10/2007
'Imports chess 'definition of chesspiece: I've just incorporated these classes in the same project at the moment.
'To do
'Check for Stalemate: repeated sequence (stop computer opponents endlessly repeating)
'Implement learning
'   With 2 computer oponents, randomly generate starting values for each player, for each imperitive value (parameter), play a game;
'   use the outcome To increase/decrease the likelihood Of choosing that value again
'For computer players pause and display moves (thinking)
'Look for good things like forks and weight them (in measuring threat is the value of all threatened pieces considered?)
'game stages: Perhaps use the count of moves to adjust the weightings of certain imperitives. i.e. openness... or put an upper limit on it.
'opening: open and control
'middle: capture and support
'end game: homing in on king: aim to control king's square and all possible positions he can move to...
'look at the value of the exchange e.g. if I take this and they take me is that a good idea
'risk/benefit: for each piece: If i take this, he can take that
'   subtract the values of the pieces that threaten a piece 
'   add the values of the pieces that it threatens
'   add the number of the pieces supporting a piece
'   subtract the value of the piece and supporting pieces
'plex work out their next move subtract it, our next move add it etc.
'consider protection via being surrounded by pawns: sort of covered in restricting and reducing threat
'en passant
'promoting to other than queen
'computer castling
'checking human move
'more standard Algebraic notation

'look at whether threatened pieces are supported: done: 11/8/2008; only threatened pieces are considered

Module modChess
    Public Const MAXPIECE As Short = 32 'Total number of chess pieces in play
    Public Const BOARDTOP As Short = 0
    Public Const BOARDLEFT As Short = 0
    Public Const MaxX As Short = 8
    Public Const MaxY As Short = 8
    Public PWIDTH As Integer = 32 'pixels
    Public PHEIGHT As Integer '= 32 'pixels: width of a board picture thing for mouse calculations
    Public PIECEWIDTH As Short '= 480 'PWIDTH * 15 '32 * 15 = 480 'twips: both piece and square
    Public PIECEHEIGHT As Short '= 480 'PHEIGHT * 15 '
    Public BOARDBOTTOM As Integer '= (MaxY + 1) * PIECEHEIGHT
    Public BBOTTOM As Integer '= '(MaxY + 1) * PHEIGHT
    Public aBoard As Board
    Public aChessPiece(MAXPIECE) As ChessPiece 'NB array index must match control index on the form for the corresponding piece
    Public pieceArray(MAXPIECE) As System.Windows.Forms.PictureBox
    Public aGame As New Game
    Public Turn As Short 'Who's turn is it? Black/Blue = -1 Top going down or White/Red = 1 Bottom going up
    Public Const TOPGOINGDOWN As Short = -1 'Direction/ Turn constant
    Public Const BOTTOMGOINGUP As Short = 1
    Public PlayerisHuman(2) As Boolean 'bit ugly but make player 0 top and player 2 bottom
    Const captureMove As Byte = 2 ' e.g. aChessPiece(aPiece).getscore(aMove) = captureMove

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
                    If aChessPiece(i).xPos <> X Or aChessPiece(i).yPos <> (MaxY + 1) - Y Then
                        Stop 'checking that pieces are in the same position board thinks they're in
                    End If
                End If
                If aBoard.GetGBoardDir(X, 9 - Y) = -1 Then System.Diagnostics.Debug.Write("-") Else System.Diagnostics.Debug.Write(" ")
                System.Diagnostics.Debug.Write(i.ToString("00"))
            Next X
            System.Diagnostics.Debug.WriteLine("")
        Next Y
    End Sub

    Sub doBestMove(ByVal pbest As Byte, ByVal mbest As Byte)
        'Create move object and use it to update the board, the pieces, record the move and display the piece
        Dim aMove As New Move 'create a new move object for storing the current move in 

        aChessPiece(pbest).checkMoves() 'Updates MaxMove of that Piece (necessary to restore correct move?!)
        aMove.p1id = pbest 'record the first piece's move
        aMove.p1XOrig = aChessPiece(pbest).xPos
        aMove.p1YOrig = aChessPiece(pbest).yPos
        aMove.p1XDest = aChessPiece(pbest).getmovex(mbest)
        aMove.p1YDest = aChessPiece(pbest).getmovey(mbest)
        aMove.p2id = aBoard.GetGBoardID(aMove.p1XDest, aMove.p1YDest)
        If aMove.p2id > 0 Then 'there's a piece being taken
            aMove.p2XOrig = aMove.p1XDest
            aMove.p2YOrig = aMove.p1YDest
        End If
        If Take(aMove.p1XDest, aMove.p1YDest) Then Beep() 'update the board and taken piece if necessary
        aChessPiece(aMove.p1id).Move(aMove.p1XDest, aMove.p1YDest) 'update the moving piece's location (and the board)
        If ispromoting(aMove.p1id, aMove.p1YDest) Then 'it realises it's reached the 8th rank
            Queening(aMove.p1id) 'change into a queen
            aMove.MCode = "q" 'in case we need to undo
        End If
        aGame.add(aMove) 'add the move to the array
        My.Forms.frmBoard.ShowPiece(aMove.p1id) 'ensure the piece is visible and in the right location on the board
        My.Forms.frmBoard.StatusStrip1.Items("statFrom").Text = "From: " & aChessPiece(aMove.p1id).AlgNot & Chr(96 + aMove.p1XOrig) & aMove.p1YOrig
        My.Forms.frmBoard.StatusStrip1.Items("statTo").Text = "To: " & aChessPiece(aMove.p1id).AlgNot & Chr(96 + aMove.p1XDest) & aMove.p1YDest
        doHistory(aChessPiece(aMove.p1id).AlgNot & Chr(96 + aMove.p1XDest) & aMove.p1YDest, aChessPiece(aMove.p1id).direction)
        If isInCheck(-Turn) Then 'warn the opponent that they're in check 'just if human I guess
            'MsgBox("Check")
        End If
    End Sub

    Sub doCastle(ByVal aMove As Move)
        'carry out the castling
        aMove.p2XOrig = 5 'king's always on 5
        aMove.p2YOrig = aMove.p1YOrig 'either 1 or 8
        aMove.p1YDest = aMove.p1YOrig
        aMove.p2YDest = aMove.p1YOrig
        If aMove.p1XOrig = 1 Then 'queenside
            aMove.p1XDest = 4
            aMove.p2XDest = 3
        Else 'kingside
            aMove.p1XDest = 6
            aMove.p2XDest = 7
        End If
        aChessPiece(aMove.p1id).move(aMove.p1XDest, aMove.p1YDest) 'Move rook
        aChessPiece(aMove.p2id).move(aMove.p2XDest, aMove.p2YDest) 'Move king
        If isInCheck(Turn) Then 'Check again for legality after the move
            MsgBox("You're in check. Either you moved into it or you didn't move out of it.")
            aChessPiece(aMove.p1id).move(aMove.p1XOrig, aMove.p1YOrig) 'move piece back
            aChessPiece(aMove.p2id).move(aMove.p2XOrig, aMove.p2YOrig) 'move taken piece back
        Else
            aGame.add(aMove) 'add the move to the array
            Turn = -Turn
        End If
    End Sub

    Sub checkOpenness(ByRef aPiece As ChessPiece, ByVal currentPlayersDirection As Short, ByRef ourOpenness As Short, ByRef theirOpenness As Short)
        'Calculates the openness which (apart from pawns) equates to controlled squares
        Dim maxLegal As Short

        maxLegal = aPiece.checkMoves 'get number of legal moves
        If aPiece.Direction = currentPlayersDirection Then
            ourOpenness = ourOpenness + maxLegal  'this is also roughly the same as contolled spaces
        Else
            theirOpenness = theirOpenness + maxLegal
        End If
    End Sub

    Sub checkSupport(ByRef aPiece As ChessPiece, ByVal currentPlayersDirection As Short, ByRef ourSuppCount As Integer, ByRef theirSuppCount As Integer)
        'Find the value of all the supported pieces
        Const takePieceValPwr As Byte = 1 'this adds the capture piece value to this power
        Dim maxSupport As Byte
        Dim aMove As Byte 'each move
        Dim aDestX As Byte 'destination x
        Dim aDestY As Byte 'destination y
        Dim aSupportedID As Byte 'id of threatened piece
        Dim aValue As Byte

        'If aGame.ShareThoughts Then
        '    If aPiece.Watch Then
        '        If aPiece.xPos = 1 Then
        '            Cout("s" & ourSuppCount & " " & theirSuppCount & vbNewLine)
        '        End If
        '    End If
        'End If
        maxSupport = aPiece.checkSupport 'get an array of supporting moves
        For aMove = 1 To maxSupport 'each supporting possibility
            aDestX = aPiece.GetSuppX(aMove) 'determine the co-ordinates of the supported piece
            aDestY = aPiece.GetSuppY(aMove)
            aSupportedID = aBoard.GetGBoardID(aDestX, aDestY) 'get the id of the supported piece
            aValue = aChessPiece(aSupportedID).Value * takePieceValPwr 'get value of supported piece
            aChessPiece(aSupportedID).Supported = aChessPiece(aSupportedID).Supported + 1 'used in checkUnsupported
            If aChessPiece(aSupportedID).Threatened = 0 Then aValue = 0
            If aChessPiece(aSupportedID).AlgNot = "K" Then aValue = 0 'no point supporting the king
            If aPiece.Direction = currentPlayersDirection Then 'it's our piece
                ourSuppCount = ourSuppCount + aValue * takePieceValPwr 'Find the total level of support for other pieces
            Else 'it's their (the opponent's) piece
                theirSuppCount = theirSuppCount + aValue * takePieceValPwr 'beTakenCount = beTakenCount + aValue * AvoidingThreat 'add the value of the piece it can take to the opponents take count
            End If
        Next aMove
    End Sub

    Sub checkThreats(ByRef aPiece As ChessPiece, ByVal currentPlayersDirection As Short, ByRef takeCount As Integer, ByRef betakenCount As Integer)
        'Calculates the total value of all threatened pieces for each player
        Const takePieceValPwr As Byte = 1 'this adds the capture piece value to this power
        Const advantageousExchangeWt As Byte = 4
        Dim aMove As Byte 'each move
        Dim aDestX As Byte 'destination x
        Dim aDestY As Byte 'destination y
        Dim aThreatenedID As Byte 'id of threatened piece
        Dim aValue As Byte

        Try
            'If aGame.ShareThoughts Then
            '    If aPiece.Watch Then
            '        If aPiece.xPos = 1 Then
            '            Cout("t" & takeCount & " " & betakenCount & vbNewLine)
            '        End If
            '    End If
            'End If
            For aMove = 1 To aPiece.checkMoves 'maxLegal 'each legal move
                If aPiece.GetScore(aMove) = captureMove Then 'find the value of threatened piece
                    aDestX = aPiece.GetMoveX(aMove) 'determine the co-ordinates of the destination
                    aDestY = aPiece.GetMoveY(aMove)
                    aThreatenedID = aBoard.GetGBoardID(aDestX, aDestY) 'get the id of the destination piece
                    aValue = aChessPiece(aThreatenedID).Value * takePieceValPwr 'weight the value of threatened piece
                    If aPiece.Direction = currentPlayersDirection Then 'this piece belongs to the player whose turn it is
                        If aPiece.Value < aChessPiece(aThreatenedID).value Then 'weight advantageous exchanges
                            takeCount = takeCount + aValue  'add the value of the piece it can take to takeCount
                        Else
                            takeCount = takeCount + aValue * advantageousExchangeWt
                        End If
                    Else 'it's their (the opponent's) piece
                        aPiece.Threatened = aPiece.Threatened + 1 'keep track of how many attackers there are
                        betakenCount = betakenCount + aValue 'add the value of the piece it can take to the opponents take count
                    End If
                End If
            Next aMove
        Catch ex As Exception
            MsgBox("checkThreats: " & ex.Message)
        End Try
    End Sub

    Sub checkUnsupported(ByVal currentPlayersDirection As Short, ByRef ourUnSuppValue As Integer, ByRef theirUnSuppValue As Integer)
        '----------- works out the total value & level of unsupported pieces
        'probably should use threatened unsupported pieces
        'The loop relating to support needs to have completed for all pieces
        Const takePieceValPwr As Byte = 1.5 'this adds the capture piece value to this power
        Dim aPiece As Byte

        'If aGame.ShareThoughts Then
        '    If aChessPiece(aPiece).Watch Then
        '        Cout("t" & ourUnSuppValue & " " & theirUnSuppValue & vbNewLine)
        '    End If
        'End If
        For aPiece = 1 To MAXPIECE
            If aChessPiece(aPiece).onboard Then
                If aChessPiece(aPiece).direction = currentPlayersDirection Then 'just look at our pieces
                    If aChessPiece(aPiece).Supported = 0 Then 'ignore kings as they don't need supporting
                        If aChessPiece(aPiece).AlgNot <> "K" Then
                            ourUnSuppValue = ourUnSuppValue + aChessPiece(aPiece).value * takePieceValPwr
                        End If
                    End If
                Else 'thier pieces
                    If aChessPiece(aPiece).AlgNot <> "K" Then 'ignore kings as they don't need supporting
                        If aChessPiece(aPiece).Supported = 0 Then
                            theirUnSuppValue = theirUnSuppValue + aChessPiece(aPiece).value * takePieceValPwr
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Sub checkTotalValue(ByRef aPiece As ChessPiece, ByVal currentPlayersDirection As Short, ByRef ourPieceTtlValue As Integer, ByRef theirPieceTtlValue As Integer)
        'Find the total value of all remaining pieces
        Const takePieceValPwr As Byte = 2

        If aPiece.Direction = currentPlayersDirection Then 'just look at our pieces
            ourPieceTtlValue = ourPieceTtlValue + aPiece.Value * takePieceValPwr
        Else 'thier pieces
            theirPieceTtlValue = theirPieceTtlValue + aPiece.Value * takePieceValPwr
        End If
    End Sub

    Sub doHistory(ByVal aDest As String, ByVal aDir As Short)
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

    Sub getBestMove(ByRef pBest As Byte, ByRef mBest As Byte)
        'Basic Move weightings are used in conjunction with Board/Position strength assessments.
        Const checkIncentive As Byte = 1
        Const takeIncentive As Byte = 2 'this adds the capture piece value to this power i.e. Value^2
        Const promoteIncentive As Byte = 10
        Dim p As Byte 'each piece
        Dim pf As Byte 'first piece
        Dim pl As Byte 'last piece
        Dim ml As Byte 'number of valid moves
        Dim m As Byte '1 to number of moves
        Dim s As Long 'score for each move
        Dim sl As Long 'maximum score
        Dim result As Byte
        Dim inCheck As Boolean

        sl = Integer.MinValue ' -10000 'best score so far
        pBest = 0 'best piece index
        mBest = 0 'best move index
        If Turn = TOPGOINGDOWN Then
            pf = 17
            pl = MAXPIECE
        Else 'BOTTOMGOINGUP
            pf = 1
            pl = 16
        End If
        For p = pf To pl 'For each piece in team
            'If aGame.ShareThoughts Then
            '    If aChessPiece(p).Watch Then
            '        Cout(aChessPiece(p).algnot & "(" & p & ") " & aChessPiece(p).xpos & " " & aChessPiece(p).ypos & vbNewLine)
            '    End If
            'End If
            If aChessPiece(p).OnBoard Then 'when a piece is taken OnBoard is set to false
                ml = aChessPiece(p).checkMoves 'count legal moves a piece can make and add them to an array of moves in the piece object
                If ml > 0 Then
                    For m = 1 To ml 'For each possible move 'The list of legal moves are stored in the piece
                        aBoard.CopyBoard() 'each board object can remember a copy of itself
                        'Try piece p's mth move; which may be capturing something
                        TakeNoShow(aChessPiece(p).GetMoveX(m), aChessPiece(p).GetMoveY(m)) 'perform capture
                        aChessPiece(p).Move(m)
                        s = Score(Turn) 'Derive a score based on the resulting board arrangement
                        If isInCheck(-Turn) Then s = s + checkIncentive 'Add some incentive to check opponent
                        inCheck = isInCheck(Turn) 'Ensure you move out of check
                        aBoard.RestoreBoard()
                        PiecesFromBoard() 'recreate piece positions from board
                        result = aChessPiece(p).checkMoves 'ensure this piece has current moves again
                        If aChessPiece(p).GetScore(m) = captureMove Then 'Add some incentive to take the opponent
                            s = s + aChessPiece(aBoard.GetGBoardID(aChessPiece(p).GetMoveX(m), aChessPiece(p).GetMoveY(m))).Value ^ takeIncentive 'get the value of the opponents piece
                        End If
                        If ispromoting(p, aChessPiece(p).GetMoveY(m)) Then s = s + promoteIncentive 'Add some incentive to queen
                        If s > sl And Not inCheck Then
                            sl = s 'Keep track of the best score
                            pBest = p 'the best piece
                            mBest = m 'the best move
                        End If
                    Next m
                End If
            End If
        Next p
    End Sub

    Function isInCheck(ByVal aDir As Short) As Boolean
        'aDir is the colour of the king that could be in check
        'Checkmoves for each opposition piece and see if it's threatening the King
        'Called form Board.vb.cmdMove_click
        Dim p As Byte 'each piece
        Dim pf As Byte 'first piece
        Dim pl As Byte 'last piece
        Dim ml As Byte 'number of valid moves
        Dim m As Byte '1 to number of moves
        Dim aX As Byte
        Dim aY As Byte
        Dim aID As Byte

        isInCheck = False
        Try
            If aDir = TOPGOINGDOWN Then 'blue: identify the range of opposition pieces
                pf = 1 'so look at red pieces
                pl = 16
            Else 'BOTTOMGOINGUP
                pf = 17
                pl = MAXPIECE
            End If
            For p = pf To pl 'For each piece in the opponent's team
                ml = aChessPiece(p).checkmoves 'find how many moves the opponent's piece can make
                If ml > 0 Then
                    For m = 1 To ml 'For each possible move
                        If aChessPiece(p).getscore(m) > 1 Then 'something can be taken
                            aX = aChessPiece(p).getmovex(m) 'determine the co-ordinates of the destination
                            aY = aChessPiece(p).getmovey(m)
                            aID = aBoard.GetGBoardID(aX, aY) 'get the id of the destination piece
                            If aChessPiece(aID).AlgNot() = "K" Then 'it's the king
                                isInCheck = True
                                'System.Diagnostics.Debug.Write("Check: " & aChessPiece(p).AlgNot & " is threatening the king.")
                            End If
                        End If
                    Next m
                End If
            Next p
        Catch ex As Exception
            MsgBox("isInCheck: " & ex.Message)
        End Try
    End Function

    Function isLegal1(ByVal aMove As Move) As Boolean
        'Initial check on the validity of the move
        isLegal1 = True
        If Not aGame.Playing Then
            isLegal1 = False 'we're not in a game
        elseIf Turn <> aChessPiece(aMove.p1id).direction Then
            isLegal1 = False 'moving wrong team: not your go
        ElseIf aMove.p1XDest > MaxX Or aMove.p1YDest > MaxY Or aMove.p1XDest < 1 Or aMove.p1YDest < 1 Then
            isLegal1 = False 'moving off board
        ElseIf aMove.p2id > 0 Then
            If aChessPiece(aMove.p1id).direction = aChessPiece(aMove.p2id).direction Then
                isLegal1 = False 'trying to take own piece
            End If
        End If
    End Function

    Function ispromoting(ByVal aPiece As Byte, ByVal aDestY As Byte) As Boolean
        'since a pawn can't retreat and always starts on the second rank it's okay to test for them being on 1 or 8
        If aChessPiece(aPiece).AlgNot = "P" Then
            If aDestY = 1 Or aDestY = 8 Then 'promoting
                ispromoting = True
            Else
                ispromoting = False
            End If
        End If
    End Function

    Sub PiecesFromBoard()
        'Having restored the board
        'update each piece based on it's position on the board
        Dim X As Byte
        Dim Y As Byte
        Dim i As Integer

        For i = 1 To 32
            aChessPiece(i).xPos = 0
            aChessPiece(i).yPos = 0
            aChessPiece(i).onboard = False
            aChessPiece(i).direction = 0
            aChessPiece(i).supported = 0 'set in scoring routine
            aChessPiece(i).threatened = 0 'set in scoring routine
        Next
        For X = 1 To MaxX 'each board position
            For Y = 1 To MaxY
                i = aBoard.GetGBoardID(X, Y) 'NB piece IDs are 1 to 32
                If i > 0 Then 'there is a piece is on it (piece i)
                    aChessPiece(i).xPos = X 'set the position of the piece
                    aChessPiece(i).yPos = Y 'set location of piece
                    aChessPiece(i).onBoard = True
                    If i > 16 Then
                        aChessPiece(i).direction = TOPGOINGDOWN
                    Else
                        aChessPiece(i).direction = BOTTOMGOINGUP
                    End If
                End If
            Next Y
        Next X
    End Sub

    Sub Queening(ByVal p As Byte)
        'Change piece p into a queen
        'Note leave out the moving and showing as that's handled elsewhere
        Dim aQueen As New Queen 'sets AlgNot and value

        aQueen.Direction = aChessPiece(p).Direction
        aQueen.ObjIndex = p
        aQueen.Move(aChessPiece(p).xPos, aChessPiece(p).yPos)
        aChessPiece(p) = aQueen 'change into a queen
        If p > 16 Then
            pieceArray(p).Image = My.Resources.queen1 'pieceArray(28).Image
        Else
            pieceArray(p).Image = My.Resources.queen 'pieceArray(4).Image
        End If
        'pieceArray(p).Visible = True
    End Sub

    Function Score(ByVal d As Short) As Long
        'Check the relative strength of the two sides
        'given the direction d is the direction of the current player
        'Returns integer.minvalue if there would be no move
        'Because very large numbers can be generated use the following value with care... They are basically pairs of characteristics from our point of view and theirs.
        Const Threatening As Byte = 1 'aggressiveness - increase our threatening
        Const AvoidingThreat As Byte = 1 'wimpishness - reduce their threatening
        Const Supporting As Byte = 1 'supporting our own pieces - increase our support
        Const Undermine As Byte = 1 'reduce their support
        Const Control As Byte = 4 'aim to control more of the board
        Const Frustrate As Byte = 1 'aim to reduce their mobility and control
        Const Keep As Byte = 2 'aim to keep as many pieces as possible
        Const Erode As Byte = 1 'aim to reduce their pieces
        Const Consolidate As Byte = 5 'aim to avoid leaving piece unsupported
        Const Isolate As Byte = 1 'aim to isolate their pieces leaving them unsupported
        Dim aPiece As Byte 'for each piece
        Dim takeCount As Integer = 0 'sum the value of all their threatened pieces from proposed position
        Dim beTakenCount As Integer = 0 'sum of the value of our threatened pieces from proposed position
        Dim ourSuppCount As Integer = 0 'sum the value of all our supported pieces from proposed position
        Dim theirSuppCount As Integer = 0
        Dim ourUnSuppValue As Integer = 0 'sum the value of all the unsupported pieces from proposed position
        Dim theirUnSuppValue As Integer = 0 'sum the value of all their unsupported pieces
        Dim ourPieceTtlValue As Integer = 0 'on board
        Dim theirPieceTtlValue As Integer = 0 'on board
        Dim ourOpenness As Short = 0 'number of possible moves (roughly the same as controlled spaces)
        Dim theirOpenness As Short = 0 'for the opponent

        Score = 0
        Try
            For aPiece = 1 To MAXPIECE 'for each piece (in it's position after a proposed move)
                If aChessPiece(aPiece).OnBoard() Then
                    checkTotalValue(aChessPiece(aPiece), d, ourPieceTtlValue, theirPieceTtlValue)
                    checkOpenness(aChessPiece(aPiece), d, ourOpenness, theirOpenness) 'find who's controlling more
                    checkThreats(aChessPiece(aPiece), d, takeCount, beTakenCount) 'find who's threatening who
                    checkSupport(aChessPiece(aPiece), d, ourSuppCount, theirSuppCount) 'find out what's supported
                End If
            Next aPiece
            checkUnsupported(d, ourUnSuppValue, theirUnSuppValue) 'uses .supported property set in checkSupport (therefore can't be in same loop)
            If aGame.ShareThoughts Then
                Cout("u" & ourUnSuppValue & " " & theirUnSuppValue & vbNewLine)
                Cout("v" & ourPieceTtlValue & " " & theirPieceTtlValue & " o" & ourOpenness & " " & theirOpenness & " t" & takeCount & " " & beTakenCount & " s" & ourSuppCount & " " & theirSuppCount & "; " & vbNewLine)
            End If
            If ourOpenness = 0 Then
                Score = Long.MinValue
            Else
                'Note: there's also some scoring above as we add the value of the pieces to the takeCount
                'and back in the routine that called this some score relating to the specific move
                Score = takeCount * Threatening - beTakenCount * AvoidingThreat
                Score = Score + ourPieceTtlValue * Keep - theirPieceTtlValue * Erode
                Score = Score + ourSuppCount * Supporting - theirSuppCount * Undermine
                Score = Score + ourOpenness * Control - theirOpenness * Frustrate
                Score = Score - ourUnSuppValue * Consolidate + theirUnSuppValue * Isolate
            End If
        Catch ex As Exception
            MsgBox("Score: " & ex.Message)
        End Try
    End Function

    Sub SetupPieces()
        Dim p As Short

        For p = 1 To MAXPIECE
            Select Case p
                Case 1
                    aChessPiece(p) = New Rook
                    aChessPiece(p).Direction = BOTTOMGOINGUP
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move(1, 1)
                    pieceArray(1) = frmBoard._piece_1
                Case 2
                    aChessPiece(p) = New Knight
                    aChessPiece(p).Direction = BOTTOMGOINGUP
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move(2, 1)
                    pieceArray(2) = frmBoard._piece_2
                Case 3
                    aChessPiece(p) = New Bishop
                    aChessPiece(p).Direction = BOTTOMGOINGUP
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move(3, 1)
                    pieceArray(3) = frmBoard._piece_3
                Case 4
                    aChessPiece(p) = New Queen
                    aChessPiece(p).Direction = BOTTOMGOINGUP
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move(4, 1)
                    pieceArray(4) = frmBoard._piece_4
                Case 5
                    aChessPiece(p) = New King
                    aChessPiece(p).Direction = BOTTOMGOINGUP
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move(5, 1)
                    pieceArray(5) = frmBoard._piece_5
                Case 6
                    aChessPiece(p) = New Bishop
                    aChessPiece(p).Direction = BOTTOMGOINGUP
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move(6, 1)
                    pieceArray(6) = frmBoard._piece_6
                Case 7
                    aChessPiece(p) = New Knight
                    aChessPiece(p).Direction = BOTTOMGOINGUP
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move(7, 1)
                    pieceArray(7) = frmBoard._piece_7
                Case 8
                    aChessPiece(p) = New Rook
                    aChessPiece(p).Direction = BOTTOMGOINGUP
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move(8, 1)
                    pieceArray(8) = frmBoard._piece_8
                Case 9 To 16
                    aChessPiece(p) = New Pawn
                    aChessPiece(p).Direction = BOTTOMGOINGUP
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move((p - 1) Mod 8 + 1, 2)
                    pieceArray(9) = frmBoard._piece_9
                    pieceArray(10) = frmBoard._piece_10
                    pieceArray(11) = frmBoard._piece_11
                    pieceArray(12) = frmBoard._piece_12
                    pieceArray(13) = frmBoard._piece_13
                    pieceArray(14) = frmBoard._piece_14
                    pieceArray(15) = frmBoard._piece_15
                    pieceArray(16) = frmBoard._piece_16
                    pieceArray(p).Image = My.Resources.pawn 'pieceArray(9).Image
                Case 17 To 24
                    aChessPiece(p) = New Pawn
                    aChessPiece(p).Direction = TOPGOINGDOWN
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move((p - 1) Mod 8 + 1, 7)
                    pieceArray(17) = frmBoard._piece_17
                    pieceArray(18) = frmBoard._piece_18
                    pieceArray(19) = frmBoard._piece_19
                    pieceArray(20) = frmBoard._piece_20
                    pieceArray(21) = frmBoard._piece_21
                    pieceArray(22) = frmBoard._piece_22
                    pieceArray(23) = frmBoard._piece_23
                    pieceArray(24) = frmBoard._piece_24
                    pieceArray(p).Image = My.Resources.pawn1 'as they might have been queened
                Case 25
                    aChessPiece(p) = New Rook
                    aChessPiece(p).Direction = TOPGOINGDOWN
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move(1, 8)
                    pieceArray(25) = frmBoard._piece_25
                Case 26
                    aChessPiece(p) = New Knight
                    aChessPiece(p).Direction = TOPGOINGDOWN
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move(2, 8)
                    pieceArray(26) = frmBoard._piece_26
                      'aChessPiece(p).Watch = True
                Case 27
                    aChessPiece(p) = New Bishop
                    aChessPiece(p).Direction = TOPGOINGDOWN
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move(3, 8)
                    pieceArray(27) = frmBoard._piece_27
                Case 28
                    aChessPiece(p) = New Queen
                    aChessPiece(p).Direction = TOPGOINGDOWN
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move(4, 8)
                    pieceArray(28) = frmBoard._piece_28
                Case 29
                    aChessPiece(p) = New King
                    aChessPiece(p).Direction = TOPGOINGDOWN
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move(5, 8)
                    pieceArray(29) = frmBoard._piece_29
                Case 30
                    aChessPiece(p) = New Bishop
                    aChessPiece(p).Direction = TOPGOINGDOWN
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move(6, 8)
                    pieceArray(30) = frmBoard._piece_30
                Case 31
                    aChessPiece(p) = New Knight
                    aChessPiece(p).Direction = TOPGOINGDOWN
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move(7, 8)
                    pieceArray(31) = frmBoard._piece_31
                Case 32
                    aChessPiece(p) = New Rook
                    aChessPiece(p).Direction = TOPGOINGDOWN
                    aChessPiece(p).ObjIndex = p
                    aChessPiece(p).Move(8, 8)
                    pieceArray(32) = frmBoard._piece_32
            End Select
            My.Forms.frmBoard.ShowPiece(p)
        Next p
    End Sub

    Function Take(ByRef X As Byte, ByRef Y As Byte) As Boolean
        'Call this prior to moving.
        Dim p As Byte

        Take = False
        On Error GoTo errTake
        p = aBoard.GetGBoardID(X, Y)
        If p > 0 Then
            aChessPiece(p).Remove()
            pieceArray(p).Visible = False 'make invisible
            Beep()
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
        'Change piece p into a pawn
        'For use when undoing a promoting
        Dim aPawn As New Pawn 'sets AlgNot and value

        aPawn.Direction = aChessPiece(p).direction
        aPawn.ObjIndex = p
        aPawn.Move(aChessPiece(p).xpos, aChessPiece(p).ypos)
        aChessPiece(p) = aPawn 'change into a pawn
        If p > 16 Then 'blue
            pieceArray(p).Image = My.Resources.pawn1 'pieceArray(17).Image
        Else 'red
            pieceArray(p).Image = My.Resources.pawn 'pieceArray(9).Image
        End If
    End Sub

End Module