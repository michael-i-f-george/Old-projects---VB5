Attribute VB_Name = "modTest"
Option Explicit

Private Sub Main()
 Dim lI As Long
 Dim bFlag As Boolean
 Dim vrtBeginningTime As Date
 Dim vrtFinishingTime
   
 bFlag = True
 vrtBeginningTime = Time
 For lI = 0 To 99999999
    If bFlag Then
       bFlag = True
    End If
 Next lI
 vrtFinishingTime = Time
 MsgBox vrtBeginningTime
 MsgBox vrtFinishingTime
 MsgBox DateDiff("s", vrtBeginningTime, vrtFinishingTime)
End Sub

