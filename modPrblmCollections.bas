Attribute VB_Name = "Module1"
Option Explicit

Const PAWN As Integer = 0
Const KNIGHT As Integer = 1
Const BISHOP As Integer = 2
Const ROOK As Integer = 3
Const QUEEN As Integer = 4
Const KING As Integer = 5

Private Sub Main()
 'Array of collections is used to store the pieces.
 Dim acolWhite(KING) As New Collection
 Dim acolBlack(KING) As New Collection '"KING" is the highest constant.
 Dim aeffacer(5) As Integer
 'To browse throught the collection.
 Dim vrtChessman As Variant
 
 
 aeffacer(0) = 0
 aeffacer(1) = 1
 aeffacer(2) = 2
 aeffacer(3) = 3
 aeffacer(4) = 4
 
 bidon aeffacer()
 
 Initialisation acolWhite()
 'Display all black bishops coordinates".
 For Each vrtChessman In acolBlack(BISHOP)
    MsgBox vrtChessman
 Next
End Sub

Private Sub bidon(autrenom() As Integer)
 Dim nI As Integer

 For nI = 0 To 4
    MsgBox autrenom(nI)
 Next nI
End Sub

Private Sub Initialisation(acolWhite() As Collection)
 'Set the chessboard up.
 'Voir fonction array.
 Dim bidon As New Collection
 
 bidon.Add "a2"
 bidon.Add "b2"
 bidon.Add "c2"
 Set acolWhite(PAWN) = bidon
 
 acolWhite(PAWN).Add "a2"
 acolWhite(PAWN).Add "a2"
 acolWhite(PAWN).Add "b2"
 acolWhite(PAWN).Add "c2"
 acolWhite(PAWN).Add "d2"
 acolWhite(PAWN).Add "e2"
 acolWhite(PAWN).Add "f2"
 acolWhite(PAWN).Add "g2"
 acolWhite(PAWN).Add "h2"
 acolWhite(KNIGHT).Add "b1"
 acolWhite(KNIGHT).Add "g1"
 acolWhite(BISHOP).Add "c1"
 acolWhite(BISHOP).Add "f1"
 acolWhite(ROOK).Add "a1"
 acolWhite(ROOK).Add "h1"
 acolWhite(QUEEN).Add "d1"
 acolWhite(KING).Add "e1"
'Idem for black men.
' acolBlack(PAWN).Add "a7"
' acolBlack(PAWN).Add "b7"
' acolBlack(PAWN).Add "c7"
' acolBlack(PAWN).Add "d7"
' acolBlack(PAWN).Add "e7"
' acolBlack(PAWN).Add "f7"
' acolBlack(PAWN).Add "g7"
' acolBlack(PAWN).Add "h7"
' acolBlack(KNIGHT).Add "b8"
' acolBlack(KNIGHT).Add "g8"
' acolBlack(BISHOP).Add "c8"
' acolBlack(BISHOP).Add "f8"
' acolBlack(ROOK).Add "a8"
' acolBlack(ROOK).Add "h8"
' acolBlack(QUEEN).Add "d8"
' acolBlack(KING).Add "e8"
End Sub
