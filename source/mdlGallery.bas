Attribute VB_Name = "mdlGallery"
'@Folder("ribbon")
Option Explicit

Sub SetStyle_onAction(control As IRibbonControl, itemId As String, index As Integer)
    Select Case itemId
    Case "title1"
        mdlStyle.SetTitle1
    Case "title2"
        mdlStyle.SetTitle2
    Case "title3"
        mdlStyle.SetTitle3
    Case "title4"
        mdlStyle.SetTitle4
    Case "title5"
        mdlStyle.SetTitle5
    Case "body1"
        mdlStyle.SetBody1
    Case "body2"
        mdlStyle.SetBody2
    Case "body3"
        mdlStyle.SetBody3
    Case "body4"
        mdlStyle.SetBody4
    Case "body5"
        mdlStyle.SetBody5
    End Select
End Sub

