Public Class Player

    Private ReadOnly d As Short 'direction
    Private ReadOnly Gradients(16) As Double

    Public Sub New(d As Short)
        Me.d = d
        For a As Integer = 1 To 16
            Gradients(a) = 1
        Next
    End Sub

    Friend Property CheckIncentive As Double = 1

    Friend Property TakeIncentive As Double = 2 'this adds the capture Piece value to this power i.e. Value^2

    Friend Property PromoteIncentive As Double = 10

    Friend Property DeterRepetition As Double = 10

    Friend Property CaptureMove As Double = 2 ' e.g. aChessPiece(aPiece).getscore(aMove) = captureMove

    Friend Property Threatening As Double = 0 'aggressiveness - increase our threatening

    Friend Property AvoidingThreat As Double = 2 'wimpishness - reduce their threatening

    Friend Property Supporting As Double = 1 'supporting our own Pieces - increase our support

    Friend Property Undermine As Double = 1

    Friend Property Control As Double = 1 'aim to control more of the board

    Friend Property Frustrate As Double = 1 'aim to reduce their mobility and control

    Friend Property Keep As Double = 1 'aim to keep as many Pieces as possible

    Friend Property Erode As Double = 1 'aim to reduce their Pieces

    Friend Property Consolidate As Double = 8 'aim to avoid leaving Piece unsupported

    Friend Property Isolate As Double = 1 'aim to isolate their Pieces leaving them unsupported

    Friend Property CheckMateIncentive As Double = 20

    Friend Function GetGradient(attribute As Integer)
        GetGradient = Gradients(attribute)
    End Function

    Friend Sub SetChangeFactor(attribute As Integer, gValue As Double)
        Gradients(attribute) = gValue
    End Sub

    Friend Sub GetBestMove(ByRef pBest As Byte, ByRef mBest As Byte)
        'Basic Move weightings are used in conjunction with Board/Position strength assessments.
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
                If ml > 0 Then
                    For m = 1 To ml 'For each possible move 'The list of legal moves are stored in the Piece
                        aBoard.CopyBoard() 'each board object can remember a copy of itself
                        'Try Piece p's mth move; which may be capturing something
                        TakeNoShow(aChessPiece(p).GetMoveX(m), aChessPiece(p).GetMoveY(m)) 'perform capture
                        aChessPiece(p).Move(m)
                        s = Score(Turn) 'Derive a score based on the resulting board arrangement
                        If IsInCheck(-Turn) Then s += CheckIncentive 'Add some incentive to check opponent
                        If IsCheckMate(-Turn) Then s += CheckMateIncentive 'Add some incentive to check opponent
                        inCheck = IsInCheck(Turn) 'Ensure you move out of check
                        aBoard.RestoreBoard()
                        PiecesFromBoard() 'recreate Piece positions from board
                        result = aChessPiece(p).CheckMoves 'ensure this Piece has current moves again
                        If aChessPiece(p).GetScore(m) = CaptureMove Then 'Add some incentive to take the opponent
                            s += aChessPiece(aBoard.GetGBoardID(aChessPiece(p).GetMoveX(m), aChessPiece(p).GetMoveY(m))).Value ^ TakeIncentive 'get the value of the opponents Piece
                            If My.Settings.ShareThoughts And aChessPiece(p).Watch Then
                                Thoughts += "Capture move: " & aChessPiece(aBoard.GetGBoardID(aChessPiece(p).GetMoveX(m), aChessPiece(p).GetMoveY(m))).Value ^ TakeIncentive & vbNewLine
                            End If
                        End If
                        If Ispromoting(p, aChessPiece(p).GetMoveY(m)) Then
                            s += PromoteIncentive 'Add some incentive to queen
                            If My.Settings.ShareThoughts And aChessPiece(p).Watch Then
                                Thoughts += "Promote incentive added: " & vbNewLine
                            End If
                        End If
                        repetitions = CountRepetition(p, m)
                        s -= DeterRepetition * repetitions
                        If s > sl And Not inCheck And repetitions < 3 Then
                            sl = s 'Keep track of the best score
                            pBest = p 'the best Piece
                            mBest = m 'the best move
                            If My.Settings.ShareThoughts Then
                                Cout(vbNewLine & "Piece: " & aChessPiece(p).AlgNot & Chr(64 + aChessPiece(p).GetMoveX(m)) & "" & aChessPiece(p).GetMoveY(m) & " Score: " & s & "--> ")
                                If aChessPiece(p).Watch Then
                                    Cout(Thoughts)
                                End If
                            End If
                        End If
                    Next m
                End If
            End If
        Next p
    End Sub

    Private Function Score(Turn As Short) As Double
        'Check the relative strength of the two sides
        'given the direction d is the direction of the current player
        'Returns integer.minvalue if there would be no move
        'Because very large numbers can be generated use the following value with care... They are basically pairs of characteristics from our point of view and theirs.

        Dim aPiece As Double 'for each Piece
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
                If aChessPiece(aPiece).OnBoard() Then 'The second 2 values are changing
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

    Friend Function GetRawScore() As Double
        Dim ourPieceTtlValue As Integer = 0 'on board
        Dim theirPieceTtlValue As Integer = 0 'on board
        Dim ourOpenness As Short = 0 'number of possible moves (roughly the same as controlled spaces)
        Dim theirOpenness As Short = 0 'for the opponent
        Dim takeCount As Integer = 0 'sum the value of all their threatened Pieces from proposed position
        Dim beTakenCount As Integer = 0 'sum of the value of our threatened Pieces from proposed position
        Dim ourSuppCount As Integer = 0 'sum the value of all our supported Pieces from proposed position
        Dim theirSuppCount As Integer = 0
        Dim ourUnSuppValue As Integer = 0 'sum the value of all the unsupported Pieces from proposed position
        Dim theirUnSuppValue As Integer = 0 'sum the value of all their unsupported Pieces

        For aPiece = 1 To MAXPIECE 'for each Piece (in it's position after a proposed move)
            If aChessPiece(aPiece).OnBoard() Then 'The second 2 values are changing
                CheckTotalValue(aChessPiece(aPiece), d, ourPieceTtlValue, theirPieceTtlValue)
                CheckOpenness(aChessPiece(aPiece), d, ourOpenness, theirOpenness) 'find who's controlling more
                CheckThreats(aChessPiece(aPiece), d, takeCount, beTakenCount) 'find who's threatening who
                CheckSupport(aChessPiece(aPiece), d, ourSuppCount, theirSuppCount) 'find out what's supported
            End If
        Next aPiece
        CheckUnsupported(d, ourUnSuppValue, theirUnSuppValue) 'uses .supported property set in checkSupport (therefore can't be in same loop)
        GetRawScore = (ourPieceTtlValue - theirPieceTtlValue)
        GetRawScore += (ourOpenness - theirOpenness)
        GetRawScore += (takeCount - beTakenCount)
        GetRawScore += (ourSuppCount - theirSuppCount)
        GetRawScore += (-ourUnSuppValue + theirUnSuppValue)
    End Function

    Friend Function ChangeAttribute(attribute As Integer, gradient As Double) As Double
        Dim x1 As Double
        Dim x2 As Double

        x1 = GetAttributeValue(attribute)
        x2 = x1 + 0.1 * gradient
        SetAttributeValue(attribute, x2)
        ChangeAttribute = x2
    End Function

    Private Sub SetAttributeValue(attribute As Integer, x2 As Double)
        Select Case attribute
            Case 1 'CheckIncentive
                CheckIncentive = x2
            Case 2 'Friend Property TakeIncentive As Double = 2 'this adds the capture Piece value to this power i.e. Value^2
                TakeIncentive = x2
            Case 3 'Friend Property PromoteIncentive As Double = 10
                PromoteIncentive = x2
            Case 4 'Friend Property DeterRepetition As Double = 10
                DeterRepetition = x2
            Case 5 'Friend Property CaptureMove As Double = 2 ' e.g. aChessPiece(aPiece).getscore(aMove) = captureMove
                CaptureMove = x2
            Case 6 'Friend Property Threatening As Double = 0 'aggressiveness - increase our threatening
                Threatening = x2
            Case 7 'Friend Property AvoidingThreat As Double = 2 'wimpishness - reduce their threatening
                AvoidingThreat = x2
            Case 8 'Friend Property Supporting As Double = 1 'supporting our own Pieces - increase our support
                Supporting = x2
            Case 9 'Friend Property Undermine As Double = 1
                Undermine = x2
            Case 10 'Friend Property Control As Double = 1 'aim to control more of the board
                Control = x2
            Case 11 'Friend Property Frustrate As Double = 1 'aim to reduce their mobility and control
                Frustrate = x2
            Case 12 'Friend Property Keep As Double = 1 'aim to keep as many Pieces as possible
                Keep = x2
            Case 13 'Friend Property Erode As Double = 1 'aim to reduce their Pieces
                Erode = x2
            Case 14 'Friend Property Consolidate As Double = 8 'aim to avoid leaving Piece unsupported
                Consolidate = x2
            Case 15 'Friend Property Isolate As Double = 1 'aim to isolate their Pieces leaving them unsupported
                Isolate = x2
            Case 16 'Friend Property Isolate As Double = 1 'aim to isolate their Pieces leaving them unsupported
                CheckMateIncentive = x2
        End Select
    End Sub

    Friend Function GetAttributeValue(attribute As Integer) As Double
        GetAttributeValue = 0
        Select Case attribute
            Case 1 'CheckIncentive
                GetAttributeValue = CheckIncentive
            Case 2 'Friend Property TakeIncentive As Double = 2 'this adds the capture Piece value to this power i.e. Value^2
                GetAttributeValue = TakeIncentive
            Case 3 'Friend Property PromoteIncentive As Double = 10
                GetAttributeValue = PromoteIncentive
            Case 4 'Friend Property DeterRepetition As Double = 10
                GetAttributeValue = DeterRepetition
            Case 5 'Friend Property CaptureMove As Double = 2 ' e.g. aChessPiece(aPiece).getscore(aMove) = captureMove
                GetAttributeValue = CaptureMove
            Case 6 'Friend Property Threatening As Double = 0 'aggressiveness - increase our threatening
                GetAttributeValue = Threatening
            Case 7 'Friend Property AvoidingThreat As Double = 2 'wimpishness - reduce their threatening
                GetAttributeValue = AvoidingThreat
            Case 8 'Friend Property Supporting As Double = 1 'supporting our own Pieces - increase our support
                GetAttributeValue = Supporting
            Case 9 'Friend Property Undermine As Double = 1
                GetAttributeValue = Undermine
            Case 10 'Friend Property Control As Double = 1 'aim to control more of the board
                GetAttributeValue = Control
            Case 11 'Friend Property Frustrate As Double = 1 'aim to reduce their mobility and control
                GetAttributeValue = Frustrate
            Case 12 'Friend Property Keep As Double = 1 'aim to keep as many Pieces as possible
                GetAttributeValue = Keep
            Case 13 'Friend Property Erode As Double = 1 'aim to reduce their Pieces
                GetAttributeValue = Erode
            Case 14 'Friend Property Consolidate As Double = 8 'aim to avoid leaving Piece unsupported
                GetAttributeValue = Consolidate
            Case 15 'Friend Property Isolate As Double = 1 'aim to isolate their Pieces leaving them unsupported
                GetAttributeValue = Isolate
            Case 16 'Friend Property Isolate As Double = 1 'aim to isolate their Pieces leaving them unsupported
                GetAttributeValue = CheckMateIncentive
        End Select
    End Function
End Class
