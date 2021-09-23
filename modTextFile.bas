Attribute VB_Name = "Module1"
'Reads a text file word by word, 13 apr 2000, Michaël George.
Option Explicit

Private Const PATHNAME As String = "d:\pgn\sample.pgn"

Private Sub Main()
   ReadFile
End Sub

Private Sub ReadFile()
   Dim strLine As String
   Dim nSampleFile As Integer
   
   nSampleFile = FreeFile
   Open PATHNAME For Input As #nSampleFile
   While Not EOF(nSampleFile)
      Line Input #nSampleFile, strLine
      SplitLine (strLine)
   Wend
   Close #nSampleFile
End Sub

Private Sub SplitLine(strLine As String)
   Dim lStart As Long
   Dim lEnd As Long
   
   lStart = 1
   lEnd = InStr(lStart + 1, strLine, " ")
   While lEnd <> 0
      MsgBox Mid$(strLine, lStart, lEnd - lStart)
      lStart = lEnd
      lEnd = InStr(lEnd + 1, strLine, " ")
   Wend
   MsgBox Mid$(strLine, lStart, Len(strLine) - lStart + 1)
End Sub
