Attribute VB_Name = "modMain"
'Aucune détection d'erreurs!
'Ne respecte pas "import format" du standard PGN (ie: plusieurs tags par
' lignes).
'Emploi de la classe "clsGame" bizarre: global?!
'Voir si on peut faire une énumération pour la nature de la pièce.
Option Explicit

Private Type CoupIntermediaire
   sNaturePiece As String * 1
   sColonneDepart As String * 1
   nRangeeDepart As Integer
   sColonneArrivee As String * 1
   nRangeeArrivee As Integer
   sPiecePromotion As String * 1
End Type

Const FICHIER As String = "d:\pgn\copie.pgn"

Public aGames(450) As clsGame
Public compteur As Integer
Dim game As clsGame
'Private row_content As String 'Définir en local

Sub Main()
 'Lecture du fichier
 Dim sLigne As String
 Dim sAeffacer As String
 Dim ciAeffacer As CoupIntermediaire
 
 Open FICHIER For Input As #1
 Do While Not EOF(1)
    Line Input #1, sLigne
    Traitement_Tag_Pair (sLigne)
 Loop
 Close #1
 
 frmListeParties.Show
 'sAeffacer = InputBox("Coup?")
 'ciAeffacer = CreationCoupIntermediaire(sAeffacer)
 'MsgBox "Pièce: " & ciAeffacer.sNaturePiece & ";   Départ: (" & ciAeffacer.sColonneDepart & "," & ciAeffacer.nRangeeDepart & ")" & ";   Arrivée: (" & ciAeffacer.sColonneArrivee & "," & ciAeffacer.nRangeeArrivee & ")" & ";   Promotion: " & ciAeffacer.sPiecePromotion
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
        'aGames(compteur).Eventt = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
     Case 2
        game.Site = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
        'aGames(compteur).Site = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
     Case 3
        game.Datee = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
        'aGames(compteur).Datee = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
     Case 4
        game.Round = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
        'aGames(compteur).Round = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
     Case 5
        game.White = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
        'aGames(compteur).White = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
     Case 6
        game.Black = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
        'aGames(compteur).Black = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
     Case 7
        game.Result = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
        'aGames(compteur).Result = Mid(sTagPair, nFinTagName + 2, Len(sTagPair) - 3 - nFinTagName)
        Set aGames(compteur) = game
        compteur = compteur + 1
        'games.Add game
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

Public Function CreationCoupIntermediaire(sCoup As String) As CoupIntermediaire
   'La chaîne sCoup est parcourue d'arrière en avant.
   Dim nPositionCurseur As Integer
   Dim ciTemp As CoupIntermediaire
      
   '
   '
   '!!! Eliminer le roque d'entree !!!
   'Améliorer l'initialisation
   '
   '
   
   'Initialisation
   ciTemp.sNaturePiece = "p"
   ciTemp.sColonneDepart = " "
   ciTemp.nRangeeDepart = 0
   ciTemp.sColonneArrivee = " "
   ciTemp.nRangeeArrivee = 0
   ciTemp.sPiecePromotion = " "
   
   'Traitement
   nPositionCurseur = Len(sCoup)
   While nPositionCurseur > 0
      Select Case Mid(sCoup, nPositionCurseur, 1)
         Case "a" To "h"
            MsgBox Len(ciTemp.sColonneArrivee)
            If ciTemp.sColonneArrivee = " " Then
               ciTemp.sColonneArrivee = Mid(sCoup, nPositionCurseur, 1)
            Else
               ciTemp.sColonneDepart = Mid(sCoup, nPositionCurseur, 1)
            End If
         Case "1" To "8"
            If ciTemp.nRangeeArrivee = 0 Then
               ciTemp.nRangeeArrivee = Val(Mid(sCoup, nPositionCurseur, 1))
            Else
               ciTemp.nRangeeDepart = Val(Mid(sCoup, nPositionCurseur, 1))
            End If
          Case "N", "B", "R", "Q", "K"
            If ciTemp.nRangeeArrivee > 0 Then  '<=> a-t'on déjà les coordonnées de la case d'arrivée?
               ciTemp.sNaturePiece = Mid(sCoup, nPositionCurseur, 1)
            Else
               ciTemp.sPiecePromotion = Mid(sCoup, nPositionCurseur, 1)
            End If
         '=, +, ' ', #, !, ? and x are ignored.
      End Select
      nPositionCurseur = nPositionCurseur - 1
      CreationCoupIntermediaire = ciTemp
   Wend
End Function
