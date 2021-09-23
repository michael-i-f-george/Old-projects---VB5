Attribute VB_Name = "modTagRecognition"
'Aucune détection d'erreurs!
'Ne respecte pas "import format" du standard PGN (ie: plusieurs tags par
' lignes).
Option Explicit
Const fichier As String = "d:\pgn\copie.pgn"
Public games As New Collection 'Faire passer en private!
Dim game As clsGame
'Private row_content As String 'Définir en local

Sub main()
 'Lecture du fichier
 Dim sLigne As String
 
 Open fichier For Input As #1
 Do While Not EOF(1)
    Line Input #1, sLigne
    Traitement_Tag_Pair (sLigne)
 Loop
 Close #1
' MsgBox games(5).White
' games(5).White = "Lazlo, Victor"
' MsgBox games(5).White
' Set games(5) = games(6)
' Dim parties(3) As clsGame
' Set parties(0) = games(1)
' Set parties(1) = games(2)
' Set parties(2) = games(3)
' MsgBox parties(0).White & " " & parties(1).Black & " " & parties(2).Datee
 'Affichage de la grille
 'Set aGames(5) = aGames(4)
 frmListeParties.Show
 'frmRecherche.Show
 MsgBox games.Count
End Sub

Private Sub Traitement_Tag_Pair(sTagPair As String)
 Dim sTagName As String, sRowContent As String
 Dim nTagNumber As Integer, nFinTagName As Integer
 
 nTagNumber = 0
 sTagPair = Trim(sTagPair)
 nFinTagName = InStr(1, sTagPair, " ")
 sTagName = Mid(sTagPair, 1, nFinTagName)
 If Left(sTagName, 1) = "[" Then
    sTagName = Mid(sTagName, 2, nFinTagName)
    nTagNumber = Identify_Tag(sTagName)
 End If
 Select Case nTagNumber
     Case 1
        Set game = New clsGame
        game.Eventt = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
     Case 2
        game.Site = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
     Case 3
        game.Datee = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
     Case 4
        game.Round = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
     Case 5
        game.White = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
     Case 6
        game.Black = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
     Case 7
        game.Result = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
        games.Add game
 End Select
End Sub

Private Function Identify_Tag(sMot As String) As Integer
'Fait correspondre à chaque sTagName un nTagNumber.
'Ce dernier vaut 0 si la recherche a échoué.
 Dim nI As Integer
 Dim sKnownTags(0 To 26) As String

 sKnownTags(0) = ""                'Débordement d'indice (while)
 sKnownTags(1) = "Event"
 sKnownTags(2) = "Site"
 sKnownTags(3) = "Date"
 sKnownTags(4) = "Round"
 sKnownTags(5) = "White"
 sKnownTags(6) = "Black"
 sKnownTags(7) = "Result"
 sKnownTags(8) = "WhiteTitle"
 sKnownTags(9) = "BlackTitle"
 sKnownTags(10) = "WhiteFRBE"
 sKnownTags(11) = "BlackFRBE"
 sKnownTags(12) = "WhiteCountry"
 sKnownTags(13) = "BlackCountry"
 sKnownTags(14) = "WhiteTeam"
 sKnownTags(15) = "BlackTeam"
 sKnownTags(16) = "Board"
 sKnownTags(17) = "Division"
 sKnownTags(18) = "Section"
 sKnownTags(19) = "Stage"
 sKnownTags(20) = "Mode"
 sKnownTags(21) = "Termination"
 sKnownTags(22) = "TimeControl"
 sKnownTags(23) = "ECO"
 sKnownTags(24) = "NIC"
 sKnownTags(25) = "Opening"
 sKnownTags(26) = "Variation"

 sMot = RTrim(sMot)
 nI = 26
 While (nI > 0) And (StrComp(sKnownTags(nI), sMot) <> 0)
    nI = nI - 1
 Wend
 Identify_Tag = nI
End Function
