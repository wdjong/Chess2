Option Strict Off
Option Explicit On

Public Class ChessPiece
    Private mstrAlgNot As String 'Algebraic Notation of Piece
    Private mlngDebugID As Integer 'ID of Piece should match array index of Piece objects and controls on the form
    Private mintDirection As Short '1 is up -1 is down
    Private mbytMaxMove As Byte 'Number of valid moves 1 to n
    Private mbytObjIndex As Byte 'object index to be stored on board
    Private mbooOnBoard As Boolean 'If the Piece is currently in play
    Private ReadOnly mbytValue As Byte 'value of Piece
    Private mbytXPos As Byte 'Current file 1 to 8 left to right convert to letters in algebraic notation section
    Private mbytYPos As Byte 'Current rank 1 to 8 bottom to top
    Private mSupported As Byte 'How many supporters
    Private mThreatened As Byte 'How many attackers
    Private mWatch As Boolean 'If true display the scores relating to this Piece
    Const DIRUP As Short = 1
    Const DIRDOWN As Short = -1
    Friend ReadOnly MAXPOSCOUNT As Byte = 28
    Private ReadOnly NextMove(MAXPOSCOUNT, 3) As Byte 'array of moves 1st dim: movenum, 2nd dim 0=x, 1=y, 2=score
    Private ReadOnly mSupport(8, 3) As Byte 'array of support 1st dim  movenum, 2nd dim 0=x, 1=y, 3=value
    Const XMOVE As Short = 0
    Const YMOVE As Short = 1
    Const SCOREMOVE As Short = 2
    Const SUPPVALUE As Byte = 2

    Public Overridable Property AlgNot() As String
        'the notation to use in printing the move
        Get
            AlgNot = mstrAlgNot
        End Get
        Set(ByVal value As String)
            mstrAlgNot = value
        End Set
    End Property

    Public Overridable Property DebugID() As Integer
        'This id is available in the class
        'But it's not used in the game
        Get
            DebugID = mlngDebugID
        End Get
        Set(ByVal value As Integer)
            mlngDebugID = value
        End Set
    End Property

    Public Overridable Property Direction() As Short
        'The direction of the Piece = the colour or which side it's on
        Get
            Direction = mintDirection
        End Get
        Set(ByVal Value As Short)
            mintDirection = Value '1 is going up/red/white -1 is down/blue/black
        End Set
    End Property

    Public Overridable Property MaxMove() As Byte
        'Set by checkmoves maxmove returns how many legal moves are available
        Get
            MaxMove = mbytMaxMove
        End Get
        Set(ByVal value As Byte)
            mbytMaxMove = value
        End Set
    End Property

    Public Overridable Property Supported() As Byte
        'Set in vb.score to record the count of supporting Pieces...
        Get
            Supported = mSupported
        End Get
        Set(ByVal value As Byte)
            mSupported = value
        End Set
    End Property

    Public Overridable Property Threatened() As Byte
        'Set in vb.score to record the count of attacking Pieces...
        Get
            Threatened = mThreatened
        End Get
        Set(ByVal value As Byte)
            mThreatened = value
        End Set
    End Property

    Public Overridable Property ObjIndex() As Byte
        'the this value is stored in the board object to allow us to reference a Piece on the board
        Get
            ObjIndex = mbytObjIndex
        End Get
        Set(ByVal Value As Byte)
            mbytObjIndex = Value
        End Set
    End Property

    Public Overridable Property OnBoard() As Boolean
        'Tells us if this Piece is on the board
        Get
            OnBoard = mbooOnBoard
        End Get
        Set(ByVal Value As Boolean)
            mbooOnBoard = Value
        End Set
    End Property

    Public Overridable ReadOnly Property Value() As Byte
        'Returns the value of this Piece
        Get
            Value = mbytValue
        End Get
    End Property

    Public Overridable Property Watch() As Boolean
        'whether we want to output thinking re this Piece
        Get
            Watch = mWatch
        End Get
        Set(ByVal value As Boolean)
            mWatch = value
        End Set
    End Property

    Public Overridable Property XPos() As Byte
        'returns the x value of this Piece
        Get
            XPos = mbytXPos
        End Get
        Set(ByVal Value As Byte)
            mbytXPos = Value
        End Set
    End Property

    Public Overridable Property YPos() As Byte
        'returns the y value of this Piece
        Get
            YPos = mbytYPos
        End Get
        Set(ByVal Value As Byte)
            mbytYPos = Value
        End Set
    End Property

    Overridable Function GetMoveX(ByVal n As Byte) As Byte
        GetMoveX = NextMove(n, XMOVE)
    End Function

    Overridable Sub SetMoveX(ByVal n As Byte, ByVal value As Byte)
        NextMove(n, XMOVE) = value
    End Sub

    Overridable Function GetMoveY(ByRef n As Byte) As Byte
        GetMoveY = NextMove(n, YMOVE)
    End Function

    Overridable Sub SetMoveY(ByVal n As Byte, ByVal value As Byte)
        NextMove(n, YMOVE) = value
    End Sub

    Overridable Function GetScore(ByRef n As Byte) As Byte
        GetScore = NextMove(n, SCOREMOVE)
    End Function

    Overridable Sub SetScore(ByVal n As Byte, ByVal value As Byte)
        NextMove(n, SCOREMOVE) = value
    End Sub

    Overridable Function GetSuppX(ByRef n As Byte) As Byte
        'Called from scoring routine in order ultimately to determine the value of the Piece supported
        GetSuppX = mSupport(n, XMOVE)
    End Function

    Overridable Sub SetSuppX(ByVal n As Byte, ByVal value As Byte)
        mSupport(n, XMOVE) = value
    End Sub

    Overridable Function GetSuppY(ByRef n As Byte) As Byte
        'Called from scoring routine in order ultimately to determine the value of the Piece supported
        GetSuppY = mSupport(n, YMOVE)
    End Function

    Overridable Sub SetSuppY(ByVal n As Byte, ByVal value As Byte)
        mSupport(n, YMOVE) = value
    End Sub

    Public Overridable Function Move(ByRef moveNum As Byte) As Boolean
        'with reference to a moveNum in the NextMove array (having run checkmoves)
        'carry out a move by updating the source and destination position on the board and the Piece co-ords
        On Error GoTo errMove
        Move = False 'assume error
        If moveNum > mbytMaxMove Then
            MsgBox("Piece.Move: " & moveNum & " is greater than " & mbytMaxMove)
        End If
        SetGBoard(mbytXPos, mbytYPos, 0, 0) 'remove Piece from old location
        mbooOnBoard = True
        mbytXPos = NextMove(moveNum, XMOVE) 'update Piece x location
        mbytYPos = NextMove(moveNum, YMOVE) 'update Piece y location
        Thoughts += "Move: " & Chr(63 + NextMove(moveNum, XMOVE)) & NextMove(moveNum, YMOVE) & vbNewLine
        SetGBoard(mbytXPos, mbytYPos, mbytObjIndex, mintDirection) 'place on new location
        Move = True 'success
        'printBoard
exitMove:
        Exit Function
errMove:
        MsgBox("Piece.Move: " & Err.Description)
        Resume exitMove
    End Function

    Public Overridable Function Move(ByRef X As Byte, ByRef Y As Byte) As Boolean
        'by specifying destination co-ordinates (e.g. a human move)
        'carry out the move by updating the board position values and the Piece co-ords
        Move = False 'assume error
        Try
            SetGBoard(mbytXPos, mbytYPos, 0, 0) 'clear source
            mbooOnBoard = True
            mbytXPos = X 'set new Piece xpos
            mbytYPos = Y 'set new Piece ypos
            SetGBoard(mbytXPos, mbytYPos, mbytObjIndex, mintDirection) 'set destination
            Move = True 'success
        Catch ex As Exception
            MsgBox("Move: " & ex.Message)
        End Try
    End Function

    Public Overridable Sub Remove()
        On Error GoTo errRemove
        SetGBoard(mbytXPos, mbytYPos, 0, 0) 'remove Piece from old location
        mbooOnBoard = False
        mbytXPos = 0
        mbytYPos = 0
        'mintDirection = 0
exitRemove:
        Exit Sub
errRemove:
        MsgBox(mstrAlgNot & ".Move: " & Err.Description)
        Resume exitRemove
    End Sub

    Public Overridable Function CheckMoves() As Byte
        'Add each legal move to the NextMove array of legal moves, 
        'Therefore CheckMoves must be called each time before moving using the nextmove array in a computer move
        'Each move has a score:
        ' 1 indicates a legal move 
        ' 2 score a possible taking move
        Dim posX As Short 'the destination
        Dim posY As Short 'the destination co-ord
        Dim posCount As Byte 'the count of possible moves
        Dim result As Short

        posCount = 0
        Try
            If OnBoard Then
                posX = mbytXPos
                posY = mbytYPos + Direction 'forward
                result = IsLegal(posX, posY, Direction) 'returns the same direction if you can't go there
                If result = 0 Then 'can move up 1 (must be unoccupied for a pawn to move forward)
                    posCount += 1 'a legal move found
                    If posCount >= 28 Then
                        Stop
                    End If
                    NextMove(posCount, XMOVE) = posX 'remember the co-ordinates
                    NextMove(posCount, YMOVE) = posY
                    NextMove(posCount, SCOREMOVE) = 1 'give it a modest positive score
                    If (mbytYPos = 2 And Direction = DIRUP) Or (mbytYPos = 7 And DIRDOWN) Then 'on 2nd rank
                        posY = mbytYPos + 2 * Direction 'up can move two forward initially
                        result = IsLegal(posX, posY, Direction)
                        If result = 0 Then 'can move up 2
                            posCount += 1
                            NextMove(posCount, XMOVE) = posX
                            NextMove(posCount, YMOVE) = posY
                            NextMove(posCount, SCOREMOVE) = 1
                        End If
                    End If
                End If
            End If
            mbytMaxMove = posCount
        Catch ex As Exception
            Stop
            MsgBox("CheckMoves: " & ex.Message)
        End Try
        CheckMoves = posCount
    End Function

    Public Overridable Function CheckSupport() As Byte
        'Add each support provided to the mSupport array (analogous to nextmove array)
        Dim posX As Short 'the destination
        Dim posY As Short 'the destination co-ord
        Dim posCount As Byte 'the count of possible moves
        Dim result As Short

        posCount = 0
        If OnBoard Then
            'check for taking Pieces
            posX = mbytXPos
            posY = mbytYPos + Direction 'forward
            If posX >= 1 And posX <= 8 And posY >= 1 And posY <= 8 Then 'on board
                result = IsLegal(posX, posY, Direction)
                If result = Direction Then 'there's someone on our side to support
                    posCount += 1
                    mSupport(posCount, XMOVE) = posX
                    mSupport(posCount, YMOVE) = posY
                    mSupport(posCount, SCOREMOVE) = 2
                End If
            End If
        End If
        CheckSupport = posCount
    End Function

    Public Sub New(ByVal aAlgNot As String, ByVal abytValue As Byte)
        MyBase.New()
        ' Get a debug ID number that can be returned by
        '   the read-only DebugID property.
        mstrAlgNot = aAlgNot
        mlngDebugID = GetDebugID()
        mbytValue = abytValue
        mintDirection = 0
        mbytXPos = 0
        mbytYPos = 0
        mSupported = 0 'Program will initialise
        mbooOnBoard = True
        mbytObjIndex = 0 'Program will initialise
        SetGBoard(mbytXPos, mbytYPos, mbytObjIndex, mintDirection)
        'debug.print "Initialize Piece " & DebugID
    End Sub

End Class
