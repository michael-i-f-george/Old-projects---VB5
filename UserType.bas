Attribute VB_Name = "modUserType"
Option Explicit

Private Type Adresse
    sPrenom As String
    sNom As String
    sRue As String
    lNumero As Long
    sBoite As String
    lCodePostal As Long
    sLocalite As String
End Type

Private Sub Main()
    Dim adrFamille(2) As Adresse
    Dim lI As Long
    
'Sylvie.
    adrFamille(0).sPrenom = "Sylvie"
    adrFamille(0).sNom = "Arnould"
    adrFamille(0).sRue = "rue Jules Mattez"
    adrFamille(0).lNumero = 99
    adrFamille(0).sBoite = ""
    adrFamille(0).lCodePostal = 6182
    adrFamille(0).sLocalite = "Souvret"
'Maman.
    adrFamille(1).sPrenom = "Mary-Jeanne"
    adrFamille(1).sNom = "Préaux"
    adrFamille(1).sRue = "rue Paul Pastur"
    adrFamille(1).lNumero = 24
    adrFamille(1).sBoite = ""
    adrFamille(1).lCodePostal = 7160
    adrFamille(1).sLocalite = "Chapelle-lez-Herlaimont"
'Stéphanie.
    adrFamille(2).sPrenom = "Stéphanie"
    adrFamille(2).sNom = "Van Den Berge"
    adrFamille(2).sRue = "place de Frasnes"
    adrFamille(2).lNumero = 11
    adrFamille(2).sBoite = "B1"
    adrFamille(2).lCodePostal = 6210
    adrFamille(2).sLocalite = "Frasnes-lez-Gosselies"
'Affichage.
    For lI = 0 To 2
        MsgBox adrFamille(lI).sPrenom & " " & adrFamille(lI).sNom & vbCrLf & _
            adrFamille(lI).sRue & ", " & adrFamille(lI).lNumero & adrFamille(lI).sBoite & vbCrLf & _
            adrFamille(lI).lCodePostal & " " & adrFamille(lI).sLocalite
    Next lI
End Sub

