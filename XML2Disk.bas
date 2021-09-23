Attribute VB_Name = "modMain"
Option Explicit

Private Const DESTINATIONPATH As String = "aeffacer.txt"

Public Sub Main()
    XML2Disk CreateXML
    MsgBox CreateXML.xml
End Sub

Private Function CreateXML() As MSXML.IXMLDOMNode
'Creates a XML tree.
    Dim xmlDocument As MSXML.DOMDocument
    Dim xmlTag As MSXML.IXMLDOMNode
    Dim xmlAttribute As MSXML.IXMLDOMAttribute
    
'Creating the HTML structure.
    Set xmlDocument = New MSXML.DOMDocument
    xmlDocument.async = False       'Control won't return to caller until end of "Load" method.
    If xmlDocument.loadXML("<HTML/>") = False Then        'Loads a XML document from specified location.
        MsgBox "Cannot load DOM document."
    End If
    Set xmlTag = xmlDocument.documentElement.appendChild(xmlDocument.createElement("HEAD"))
    xmlTag.Text = "Hello, world."
    Set xmlTag = xmlDocument.documentElement.appendChild(xmlDocument.createElement("BODY"))
'Adding a table.
    Set xmlTag = xmlTag.appendChild(xmlDocument.createElement("TABLE"))
    Set xmlAttribute = xmlDocument.createAttribute("WIDTH")
    xmlAttribute.Value = "100%"
    xmlTag.Attributes.setNamedItem xmlAttribute
    Set xmlAttribute = xmlDocument.createAttribute("BORDER")
    xmlAttribute.Value = "1"
    xmlTag.Attributes.setNamedItem xmlAttribute
    Set xmlTag = xmlTag.appendChild(xmlDocument.createElement("TR"))
    xmlTag.appendChild xmlDocument.createElement("TD")
    xmlTag.appendChild xmlDocument.createElement("TD")
    Set xmlTag = xmlTag.appendChild(xmlDocument.createElement("TD"))
    Set xmlAttribute = xmlDocument.createAttribute("BGCOLOR")
    xmlAttribute.Value = "#99CCFF"
    xmlTag.Attributes.setNamedItem xmlAttribute
    Set xmlTag = xmlTag.appendChild(xmlDocument.createElement("TD"))
    xmlTag.Text = "Cell 4."
    Set CreateXML = xmlDocument.documentElement
'Closing.
    Set xmlAttribute = Nothing
    Set xmlTag = Nothing
    Set xmlDocument = Nothing
End Function

Private Sub XML2Disk(xmlRoot As MSXML.IXMLDOMNode)
    Dim nFile As Integer
    
    nFile = FreeFile
    Open DESTINATIONPATH For Output As #nFile
    WriteXML xmlRoot, nFile, ""
    Close #nFile
End Sub

Private Sub WriteXML(xmlNode As MSXML.IXMLDOMNode, nFile As Integer, sIndent As String)
'Writes recursively the XML tree to a file.
    Dim xmlChild As MSXML.IXMLDOMNode
    Dim xmlAttribute As MSXML.IXMLDOMAttribute
    Dim sAttributes As String
    
    If Len(xmlNode.baseName) <> 0 Then
'ATTRIBUTES.
        For Each xmlAttribute In xmlNode.Attributes
            sAttributes = sAttributes & " " & xmlAttribute.nodeName & "=""" & xmlAttribute.Value & """"
        Next
'TAG ITSELF.
        Print #nFile, sIndent & "<" & xmlNode.baseName & sAttributes & ">"
        For Each xmlChild In xmlNode.childNodes
            WriteXML xmlChild, nFile, sIndent & "   "   'Recursive call.
        Next
        Print #nFile, sIndent & "</" & xmlNode.baseName & ">"
    Else
'TAG CONTENT.
        If Not IsNull(xmlNode.nodeValue) Then
            Print #nFile, sIndent & xmlNode.nodeValue
        End If
    End If
End Sub
