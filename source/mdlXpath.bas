Attribute VB_Name = "mdlXpath"
'@Folder "xml"
Option Explicit

Property Get TargetNode(targetTag As String, Optional termsTag As String = "", _
Optional termsText As String = "") As String
    If termsTag <> "" Then
        TargetNode = "//" + targetTag + "[descendant::" + termsTag + "='" + termsText + "']"
    Else
        TargetNode = "//" + targetTag
    End If
End Property


