Attribute VB_Name = "mdlXml"
'@Folder "xml"
Option Explicit


Property Get SelectedTexts(contextNodeTag As String, keyNodeTag As String, keyNodeText As String, _
valueNodeTag As String) As Collection

    #If RELEASE Then
        Dim nodes As Object
        Dim node As Object
    #Else
        Dim nodes As IXMLDOMNodeList
        Dim node As IXMLDOMNode
    #End If

    Set nodes = SelectedNodes(mdlXpath.TargetNode(contextNodeTag, keyNodeTag, keyNodeText))

    Dim text As String
    Dim texts As New Collection
    For Each node In nodes
        texts.Add node.SelectSingleNode(valueNodeTag).text
    Next
    Set SelectedTexts = texts
End Property

Sub AddNode(contextNodeTag As String, dicElements As Object)
        
    #If RELEASE Then
        Dim domDoc As Object
        Set domDoc = CreateObject("MSXML2.DOMDocument.6.0")
        Dim parentNode As Object
        Dim childNode As Object
    #Else
        Dim domDoc As MSXML2.DOMDocument60
        Set domDoc = New MSXML2.DOMDocument60
        Dim parentNode As IXMLDOMNode
        Dim childNode As IXMLDOMNode
    #End If
            
    Set parentNode = SelectedNode("//config")
    Set domDoc = parentNode.OwnerDocument
    
    Set parentNode = parentNode.appendChild(domDoc.createNode(NODE_ELEMENT, contextNodeTag, ""))
        
    Dim tag As Variant
    For Each tag In dicElements
        parentNode.appendChild(domDoc.createNode(NODE_ELEMENT, tag, "")).text = dicElements.Item(tag)
    Next
    
    SaveXmlDocument domDoc
    
End Sub

Sub RemoveNode(contextNodeTag As String, keyNodeTag As String, keyNodeText As String)

    #If RELEASE Then
        Dim domDoc As Object
        Set domDoc = CreateObject("MSXML2.DOMDocument.6.0")
        Dim nodes As Object
        Dim node As Object
    #Else
        Dim domDoc As MSXML2.DOMDocument60
        Set domDoc = New MSXML2.DOMDocument60
        Dim nodes As IXMLDOMNodeList
        Dim node As IXMLDOMNode
    #End If
    
    Set nodes = SelectedNodes(mdlXpath.TargetNode(contextNodeTag, keyNodeTag, keyNodeText))
    
    If nodes.Length = 0 Then
        Exit Sub
    End If
    
    For Each node In nodes
        node.parentNode.RemoveChild node
        Set domDoc = node.OwnerDocument
    Next
    
    SaveXmlDocument domDoc
    
End Sub

Private Property Get SelectedNode(xPath As String) As Object

    #If RELEASE Then
        Dim domDoc As Object
        Set domDoc = CreateObject("MSXML2.DOMDocument.6.0")
    #Else
        Dim domDoc As MSXML2.DOMDocument60
        Set domDoc = New MSXML2.DOMDocument60
    #End If
    Set domDoc = XmlDocument
    
    Set SelectedNode = domDoc.SelectSingleNode(xPath)
    
    Set domDoc = Nothing
End Property

Private Property Get SelectedNodes(xPath As String) As Object

    #If RELEASE Then
        Dim domDoc As Object
        Set domDoc = CreateObject("MSXML2.DOMDocument.6.0")
    #Else
        Dim domDoc As MSXML2.DOMDocument60
        Set domDoc = New MSXML2.DOMDocument60
    #End If
    Set domDoc = XmlDocument
    
    Set SelectedNodes = domDoc.SelectNodes(xPath)
    
    Set domDoc = Nothing
End Property

Private Sub SaveXmlDocument(domDoc As Object)
    
        '整形
    #If RELEASE Then
        Dim xmlWriter As Object
        Set xmlWriter = CreateObject("Msxml2.MXXMLWriter.6.0")
        Dim xmlReader As Object
        Set xmlReader = CreateObject("MSXML2.SAXXMLReader.6.0")
    #Else
        Dim xmlWriter As MSXML2.MXXMLWriter60
        Set xmlWriter = New MSXML2.MXXMLWriter60
        Dim xmlReader As MSXML2.SAXXMLReader60
        Set xmlReader = New MSXML2.SAXXMLReader60
    #End If
    
    
    xmlWriter.Indent = True
    Set xmlReader.contentHandler = xmlWriter
    
    xmlReader.parse domDoc
    domDoc.LoadXML xmlWriter.output
    
    'XMLファイル出力
    domDoc.Save ThisDocument.Path + "\" + CONFIG_FILE_NAME

End Sub

Private Sub CreateXML()
    #If RELEASE Then
        Dim domDoc As Object
        Set domDoc = CreateObject("MSXML2.DOMDocument.6.0")
    #Else
        Dim domDoc As MSXML2.DOMDocument60
        Set domDoc = New MSXML2.DOMDocument60
    #End If

    domDoc.async = False
    
    #If RELEASE Then
        Dim rootNode As Object
        Dim parentNode As Object
        Dim childNode1 As Object
        Dim childNode2 As Object
    #Else
        Dim rootNode As MSXML2.IXMLDOMNode
        Dim parentNode As MSXML2.IXMLDOMNode
        Dim childNode1 As MSXML2.IXMLDOMNode
        Dim childNode2 As MSXML2.IXMLDOMNode
    #End If
    Set rootNode = domDoc.appendChild(domDoc.createNode(NODE_ELEMENT, "config", ""))
    
    Set parentNode = rootNode.appendChild(domDoc.createNode(NODE_ELEMENT, TAG_KEYBINDING, ""))
    Set childNode1 = parentNode.appendChild(domDoc.createNode(NODE_ELEMENT, TAG_KEYCODE, ""))
    Set childNode2 = parentNode.appendChild(domDoc.createNode(NODE_ELEMENT, TAG_FORMER_COMMAND, ""))
    
    SaveXmlDocument domDoc
    
    Set domDoc = Nothing
End Sub

Private Property Get XmlDocument() As Object
    
    #If RELEASE Then
        Dim domDoc As Object
        Set domDoc = CreateObject("MSXML2.DOMDocument.6.0")
    #Else
        Dim domDoc As MSXML2.DOMDocument60
        Set domDoc = New MSXML2.DOMDocument60
    #End If
    
    domDoc.async = False
    
    If Dir(ThisDocument.Path + "\" + CONFIG_FILE_NAME) = "" Then
        CreateXML
    End If
    
    domDoc.Load ThisDocument.Path + "\" + CONFIG_FILE_NAME
    
    Set XmlDocument = domDoc

    Set domDoc = Nothing
End Property


