Attribute VB_Name = "modArrayArgument"
Option Explicit

Private Sub subQuelconque(nVecteur)
 Dim nI As Integer
 
 For nI = 2 To 5
    MsgBox nVecteur(nI)
 Next nI

End Sub

Private Sub Main()
 Dim nI As Integer
 Dim nVecteur(5) As Integer
   
 For nI = 0 To 5
    nVecteur(nI) = nI
 Next nI
 subQuelconque (nVecteur)
End Sub
