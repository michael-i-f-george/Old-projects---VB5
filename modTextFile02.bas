Attribute VB_Name = "modTextFile01"
'Reads a text file word by word, 13 apr 2000, Michaël George.
Option Explicit

Private Const PATHNAME As String = "d:\pgn\sample3.pgn"

Private strConcatenated As String

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
      SplitLine3 (strLine)
   Wend
   Close #nSampleFile
   'MsgBox strConcatenated
End Sub

Private Sub SplitLine3(strLine As String)
'Seems to work very well.
'Instead of testing each time if byline = 0, a step of "2" seems usable.
   Dim lStart As Long
   Dim lEnd As Long
   Dim strWord As String
   Dim byLine() As Byte
   Dim nI As Integer
   
   nI = 0
   byLine = strLine
   While nI < UBound(byLine)
      If byLine(nI) <> 0 Then
         If byLine(nI) <> vbKeySpace Then
            strWord = strWord & Chr$(byLine(nI))
         Else
            If Len(strWord) <> 0 Then MsgBox ">" & strWord & "<"
            strWord = ""
         End If
      End If
      nI = nI + 1
   Wend
   If Len(strWord) <> 0 Then MsgBox ">" & strWord & "<"
 End Sub

Private Sub SplitLine2(strLine As String)
   Dim lStart As Long
   Dim lEnd As Long
   Dim strWord As String
   Dim byLine() As Byte
   Dim nI As Integer
   
   nI = 0
   byLine = strLine
   While nI < UBound(byLine)
      While nI < UBound(byLine) And byLine(nI) <> vbKeySpace
         If byLine(nI) <> 0 Then strWord = strWord & Chr$(byLine(nI))
         nI = nI + 1
      Wend
      If Len(strWord) <> 0 Then MsgBox ">" & strWord & "<"
      strWord = ""
      While nI < UBound(byLine) And byLine(nI) = vbKeySpace
         nI = nI + 1
      Wend
    Wend
'   lStart = 1
'   lEnd = InStr(lStart + 1, strLine, " ")
'   While lEnd <> 0
'      strWord = Mid$(strLine, lStart, lEnd - lStart)
'      If Len(strWord) > 0 Then strConcatenated = strConcatenated & " >" & strWord & "<"
'      lStart = lEnd + 1
'      lEnd = InStr(lEnd + 1, strLine, " ")
'   Wend
'   strWord = Mid$(strLine, lStart, Len(strLine) - lStart + 1)
'   If Len(strWord) <> 0 Then strConcatenated = strConcatenated & " >" & strWord & "<"
End Sub

Private Sub SplitLine(strLine As String)
   Dim lStart As Long
   Dim lEnd As Long
   Dim strWord As String
   
   lStart = 1
   lEnd = InStr(lStart + 1, strLine, " ")
   While lEnd <> 0
      strWord = Mid$(strLine, lStart, lEnd - lStart)
      If Len(strWord) > 0 Then strConcatenated = strConcatenated & " >" & strWord & "<"
      lStart = lEnd + 1
      lEnd = InStr(lEnd + 1, strLine, " ")
   Wend
   strWord = Mid$(strLine, lStart, Len(strLine) - lStart + 1)
   If Len(strWord) <> 0 Then strConcatenated = strConcatenated & " >" & strWord & "<"
End Sub
