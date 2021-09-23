VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
 Dim colAeffacer As New Collection
 Dim vrtBrowser As Variant
 
 Dim a As Variant
 
 a = Array(11, 22, 33)
 MsgBox a(2)
 colAeffacer.Add Array(7, 13)
 For Each vrtBrowser In colAeffacer
    MsgBox vrtBrowser(1)
 Next
End Sub
