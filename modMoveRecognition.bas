Attribute VB_Name = "modMoveRecognition02"
Option Explicit

'Player having to move.
Private Const WHITETOMOVE As Boolean = False
Private Const BLACKTOMOVE As Boolean = True
'Chessboard colour occupation.
Private Const EMPTYSQUARE As Integer = 0
Private Const WHITEMAN As Integer = 1
Private Const BLACKMAN As Integer = 2
'Kinds of chessman.
Private Const PAWN As Integer = 0
Private Const KNIGHT As Integer = 1
Private Const BISHOP As Integer = 2
Private Const ROOK As Integer = 3
Private Const QUEEN As Integer = 4
Private Const KING As Integer = 5
'Case where the en passant file doesn't existing.
Private Const NOEPFILE As Integer = 8
'
Private Const KNIGHTCOLLECTION As Integer = 0
Private Const QUEENCOLLECTION As Integer = 0

'En passant file.
Private nEPFile As Integer
'Abilities to castle.
Private blnWhiteShortCastle As Boolean
Private blnWhiteLongCastle As Boolean
Private blnBlackShortCastle As Boolean
Private blnBlackLongCastle As Boolean
'Chessboard.
Private Const OCCUPATION As Byte = 0
Private Const PIECETYPE As Byte = 1
'anChessboard(n1, n2, OCCUPATION) = EMPTYSQUARE, WHITEMAN or BLACKMAN
'anChessboard(n1, n2, PIECETYPE)= PAWN, KNIGHT, BISHOP, ROOK, QUEEN or KING
'Private anChessboard(7, 7, 1) As Integer
Private anchessboard(7, 7) As Integer
'abyPieceBoard(n1, n2) = PAWN, KNIGHT, BISHOP, ROOK, QUEEN or KING
Private abyPiecesBoard(7, 7) As Byte
'abyOccupationboard(n1, n2) = EMPTYSQUARE, WHITEMAN or BLACKMAN
Private abyOccupationBoard(7, 7) As Byte
'private anChessboard(7,7) as Integer
Private Const BLACKKING As Integer = -6
Private Const BLACKQUEEN As Integer = -5
Private Const BLACKROOK As Integer = -4
Private Const BLACKBISHOP As Integer = -3
Private Const BLACKKNIGHT As Integer = -2
Private Const BLACKPAWN As Integer = -1
'Private Const EMPTYSQUARE As Integer = 0
Private Const WHITEPAWN As Integer = 1
Private Const WHITEKNIGHT As Integer = 2
Private Const WHITEBISHOP As Integer = 3
Private Const WHITEROOK As Integer = 4
Private Const WHITEQUEEN As Integer = 5
Private Const WHITEKING As Integer = 6
Private Const AFILE As Byte = 0
Private Const HFILE As Byte = 7

Private Sub Main()
 'Two vectors of six collections are used to store the pieces.
 Dim acolWhite(KING) As Collection
 Dim acolBlack(KING) As Collection
 Dim vrtChessman As Variant 'Collection enumerator.
 Dim anChessboardOccupation(7, 7) As Integer
 
 InitializeChessboard acolWhite(), acolBlack(), anChessboardOccupation
 
 'This example displays all black bishops coordinates.
 For Each vrtChessman In acolBlack(BISHOP)
    DisplaySquare (vrtChessman)
 Next
 'While this one gives all white pawns square.
 For Each vrtChessman In acolWhite(PAWN)
    DisplaySquare (vrtChessman)
 Next
 'Test of the "LocateKnights" function
 Dim colAeffacer As Collection
 Dim vrtaeffacer As Variant
 
 vrtaeffacer = TranslateSquare("f3")
 MsgBox TypeName(vrtaeffacer)
 MsgBox "Premiere salve de cavaliers"
 Set colAeffacer = LocateKnights(acolWhite(KNIGHT), TranslateSquare("f3"))
  For Each vrtChessman In colAeffacer
    DisplaySquare (vrtChessman)
 Next
 MsgBox "Seconde salve de cavaliers"
 Set colAeffacer = LocateKnights(acolBlack(KNIGHT), TranslateSquare("c6"))
 For Each vrtChessman In colAeffacer
    DisplaySquare (vrtChessman)
 Next
 DisplaySquare (TranslateSquare("a1"))
End Sub

Private Sub DisplaySquare(vrtSquare As Variant)
 'For debugging purposes only.
 MsgBox Chr(vrtSquare(0) + 97) & (vrtSquare(1) + 1)
End Sub

Private Function TranslateSquare(strSquare As String) As Variant
 'For debugging purposes only.
 TranslateSquare = Array(Asc(Left(strSquare, 1)) - 97, Asc(Right(strSquare, 1)) - 49)
End Function

Private Function LocatePiece(colChessmen() As Collection, anDestinationSquare() As Integer, nManKind As Integer, nPlayingColour As Integer) As Collection
   Dim colOriginSquares As Collection
   
   Select Case nManKind
      Case PAWN
      Case KNIGHT
         colOriginSquares = LocateKnights(colChessmen(KNIGHT), anDestinationSquare())
      Case BISHOP
      Case ROOK
      Case QUEEN
      Case KING
      Case Else
         MsgBox "FindThePiece function: wrong kind of chessman."
   End Select
End Function

Private Function LocatePawns(vrtHeadingSquare As Variant, blnPlayingColor As Boolean, byEPFile As Byte) As Collection
 Dim colResult As Collection
 Dim byDestFile As Byte
 Dim byDestRank As Byte
 Dim nAllyPawn As Integer
 Dim byThirdRank As Byte
 Dim byFourthRank As Byte
 Dim byOneRankBelow As Byte
 Dim byTwoRanksBelow As Byte

 Set colResult = New Collection
 'Optimization.
 byDestFile = vrtHeadingSquare(0)
 byDestRank = vrtHeadingSquare(1)
 'Readability.
 If blnPlayingColor = WHITETOMOVE Then
    nAllyPawn = WHITEPAWN
    byThirdRank = 2
    byFourthRank = 3
    byOneRankBelow = byDestRank - 1
    byTwoRanksBelow = byDestRank - 2
 Else
    nAllyPawn = BLACKPAWN
    byThirdRank = 5
    byFourthRank = 4
    byOneRankBelow = byDestRank + 1
    byTwoRanksBelow = byDestRank + 2
 End If
 'Searching.
 If anchessboard(byDestFile, byDestRank) = EMPTYSQUARE Then
    If anchessboard(byDestFile, byOneRankBelow) = nAllyPawn Then
       'One square move.
       colResult.Add Array(byDestFile, byOneRankBelow)
    Else
       If byDestRank = byFourthRank Then
          'Two squares jump.
          If anchessboard(byDestFile, byTwoRanksBelow) = nAllyPawn Then
             If anchessboard(byDestFile, byOneRankBelow) = EMPTYSQUARE Then
                colResult.Add Array(byDestFile, byTwoRanksBelow)
             End If
          End If
       Else
          'En passant capture.
          If byDestFile = byEPFile Then
             If byDestRank = byThirdRank Then
                If byDestFile <> AFILE Then
                   If anchessboard(byDestFile - 1, byOneRankBelow) = nAllyPawn Then
                      colResult.Add Array(byDestFile - 1, byOneRankBelow)
                   End If
                End If
                If byDestFile <> HFILE Then
                   If anchessboard(byDestFile + 1, byOneRankBelow) = nAllyPawn Then
                      colResult.Add Array(byDestFile + 1, byOneRankBelow)
                   End If
                End If
             End If
          End If
       End If
    End If
 Else
    'Ordinary capture.
    If byDestFile <> AFILE Then
       If anchessboard(byDestFile - 1, byOneRankBelow) = nAllyPawn Then
          colResult.Add Array(byDestFile - 1, byOneRankBelow)
       End If
    End If
    If byDestFile <> HFILE Then
       If anchessboard(byDestFile + 1, byOneRankBelow) = nAllyPawn Then
          colResult.Add Array(byDestFile + 1, byOneRankBelow)
       End If
    End If
 End If
 Set LocatePawns = colResult
 Set colResult = Nothing
End Function

Private Function LocateKnights(colKnight As Collection, vrtDestinationSquare As Variant) As Collection
 'This function does not need to know whose turn it is.
 Dim vrtKnight As Variant 'Collection enumerator.
 Dim colResult As New Collection

 For Each vrtKnight In colKnight
    'ABS(Xorigin-Xdestination) + ABS(Yorigin-Ydestination) =? 3
    If Abs(vrtDestinationSquare(0) - vrtKnight(0)) + Abs(vrtDestinationSquare(1) - vrtKnight(1)) = 3 Then
       colResult.Add vrtKnight
    End If
 Next
 Set LocateKnights = colResult
 Set colResult = Nothing
End Function

Private Function LocateBishops(colBishop As Collection, vrtDestinationSquare As Variant) As Collection
 Dim vrtBishop As Variant 'Collection enumerator.
 Dim colResult As Collection
 Dim byDestFile As Byte
 Dim byDestRank As Byte
 Dim byCurFile As Byte
 Dim byCurRank As Byte
 Dim nHoriStep As Integer
 Dim nVertiStep As Integer
 
 byDestFile = vrtDestinationSquare(0)
 byDestRank = vrtDestinationSquare(1)
 Set colResult = New Collection
 For Each vrtBishop In colBishop
    'Is destination on Bishop's diagonal?
    If Abs(byDestFile - vrtBishop(0)) = Abs(byDestRank - vrtBishop(1)) Then
       'Is there an obstacle on the way?
       If byDestFile > vrtBishop(0) Then
          nHoriStep = -1   'Browsing from the destination.
       Else
          nHoriStep = 1
       End If
       If bydstrank > vrtBishop(1) Then
          nVertiStep = -1
       Else
          nVertiStep = 1
       End If
       byCurFile = byDestFile + nHoriStep
       byCurRank = byDestRank + nVertiStep
       While anchessboard(byCurFile, byCurRank) = EMPTYSQUARE
          byCurFile = byCurFile + nHoriStep
          byCurRank = byCurRank + nVertiStep
       Wend
       If byCurFile = vrtBishop(0) Then
          colResult.Add vrtBishop
       End If
    End If
 Next
 Set LocateBishops = colResult
 Set colResult = Nothing
End Function

Private Function LocateRooks(colRook As Collection, vrtDestinationSquare As Variant) As Collection
 Dim vrtRook As Variant   'Collection enumerator.
 Dim colResults As Collection
 Dim byDestFile As Byte
 Dim byDestRank As Byte
 Dim byCurFile As Byte
 Dim byCurRank As Byte
 Dim nStep As Integer

 byDestFile = vrtDestinationSquare(0)
 byDestRank = vrtdestinationrank(1)
 Set colResults = New Collection
 For Each vrtRook In colRook
    'Is destination on Rook's file?
    If byDestFile = vrtRook(0) Then
       If byDestRank > vrtRook(1) Then
          nStep = -1
       Else
          nStep = 1
       End If
       'Is there an obstacle on the way?
       byCurFile = byDestFile + nStep
       While anchessboard(byCurFile, byDestRank) = EMPTYSQUARE
          byCurFile = byCurFile + nStep
       Wend
       If byCurFile = vrtRook(0) Then
          colResults.Add vrtRook
       End If
    Else
       'Is destination on Rook's rank?
       If byDestRank = vrtRook(1) Then
          If byDestFile > vrtRook(1) Then
             nStep = -1
          Else
             nStep = 1
          End If
          'Is there an obstacle on the way?
          byCurRank = byDestRank + nStep
          While anchessboard(byDestFile, byCurRank) = EMPTYSQUARE
             byCurRank = byCurRank + nStep
          Wend
          If byCurRank = vrtRook(1) Then
             colResults.Add vrtRook
          End If
       End If
    End If
 Next
 LocateRooks = colResults
 Set colResults = Nothing
End Function

Private Function LocateQueens(colQueen As Collection, vrtDestinationSquare As Variant) As Collection
 Dim vrtQueen As Variant   'Collection enumerator.
 Dim colBishopResults As Collection
 Dim colRookResults As Collection

 'Queens moving like Bishops and Rooks, LocateBishops and LocateRooks
 'functions are altenatively called.
 Set colBishopResults = New Collection
 Set colRookResults = New Collection
 colBishopResults = LocateBishops(colQueen, vrtDestinationSquare)
 colRookResults = LocateRooks(colQueen, vrtDestinationSquare)
 'Concatenating the two collections.
 For Each vrtQueen In colRookResults
    colBishopResults.Add vrtQueen
 Next
 Set colRookResults = Nothing
 'Returning results.
 Set LocateQueens = colBishopResults
 Set colBishopResults = Nothing
End Function

Private Sub ResolveAmbiguity(colResults As Collection)
   MsgBox "There is an ambiguity."
End Sub

Private Sub InitializeChessboard(acolWhite() As Collection, acolBlack() As Collection, anChessboardOccupation() As Integer)
 'Set the chessboard up.
 Dim nI As Integer
 
 For nI = PAWN To KING
    Set acolWhite(nI) = New Collection
    Set acolBlack(nI) = New Collection
 Next nI
 'Setting the chessboard up.
 acolWhite(PAWN).Add Array(0, 1)
 acolWhite(PAWN).Add Array(1, 1)
 acolWhite(PAWN).Add Array(2, 1)
 acolWhite(PAWN).Add Array(3, 1)
 acolWhite(PAWN).Add Array(4, 1)
 acolWhite(PAWN).Add Array(5, 1)
 acolWhite(PAWN).Add Array(6, 1)
 acolWhite(PAWN).Add Array(7, 1)
 acolWhite(KNIGHT).Add Array(1, 0)
 acolWhite(KNIGHT).Add Array(6, 0)
 acolWhite(BISHOP).Add Array(2, 0)
 acolWhite(BISHOP).Add Array(5, 0)
 acolWhite(ROOK).Add Array(0, 0)
 acolWhite(ROOK).Add Array(7, 0)
 acolWhite(QUEEN).Add Array(3, 0)
 acolWhite(KING).Add Array(4, 0)
'Idem for black men.
 acolBlack(PAWN).Add Array(0, 6)
 acolBlack(PAWN).Add Array(1, 6)
 acolBlack(PAWN).Add Array(2, 6)
 acolBlack(PAWN).Add Array(3, 6)
 acolBlack(PAWN).Add Array(4, 6)
 acolBlack(PAWN).Add Array(5, 6)
 acolBlack(PAWN).Add Array(6, 6)
 acolBlack(PAWN).Add Array(7, 6)
 acolBlack(KNIGHT).Add Array(1, 7)
 acolBlack(KNIGHT).Add Array(6, 7)
 acolBlack(BISHOP).Add Array(2, 7)
 acolBlack(BISHOP).Add Array(5, 7)
 acolBlack(ROOK).Add Array(0, 7)
 acolBlack(ROOK).Add Array(7, 7)
 acolBlack(QUEEN).Add Array(3, 7)
 acolBlack(KING).Add Array(4, 7)
'Deals with board occupation.
 For nI = 0 To 7
    anChessboardOccupation(nI, 0) = WHITEMAN
    anChessboardOccupation(nI, 1) = WHITEMAN
    anChessboardOccupation(nI, 6) = BLACKMAN
    anChessboardOccupation(nI, 7) = BLACKMAN
 Next nI
End Sub

Private Sub Class_Initialize()
 Dim nI As Integer
 
 nEnPassantFile = NOEPFILE
 blnWhiteShortCastle = True
 blnWhiteLongCastle = True
 blnBlackShortCastle = True
 blnBlackLongCastle = True
 'Setting the board up.
 For nI = 0 To 7
    abchessboard(nI, 1, 1) = WHITEPAWN
    anchessboard(nI, 6, 1) = BLACKPAWN
 Next nI
 anchessboard(0, 0, 1) = WHITEROOK
 anchessboard(1, 0, 1) = WHITEKNIGHT
 anchessboard(2, 0, 1) = WHITEBISHOP
 anchessboard(3, 0, 1) = WHITEQUEEN
 anchessboard(4, 0, 1) = WHITEKING
 anchessboard(5, 0, 1) = WHITEBISHOP
 anchessboard(6, 0, 1) = WHITEKNIGHT
 anchessboard(7, 0, 1) = WHITEROOK
 anchessboard(0, 7, 1) = BLACKROOK
 anchessboard(1, 7, 1) = BLACKKNIGHT
 anchessboard(2, 7, 1) = BLACKBISHOP
 anchessboard(3, 7, 1) = BLACKQUEEN
 anchessboard(4, 7, 1) = BLACKKING
 anchessboard(5, 7, 1) = BLACKBISHOP
 anchessboard(6, 7, 1) = BLACKKNIGHT
 anchessboard(7, 7, 1) = BLACKROOK
End Sub

