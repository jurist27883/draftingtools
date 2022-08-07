Attribute VB_Name = "mdlVersionChecker"
'@Folder "version"
Option Explicit

Const thisVerion = "v2.4.1"
Const key = "tag_name"

Sub CheckUpdate()
    
    Dim httpReq As New XMLHTTP60                 '「Microsoft XML, v6.0」を参照設定
    Dim params As New Scripting.Dictionary       '「Microsoft Scripting Runtime」を参照設定
    

    With httpReq
        .Open "GET", "https://api.github.com/repos/sakura-editor/sakura/releases/latest"
        .send
    End With

    Do While httpReq.readyState < 4
        DoEvents
    Loop
    
    Dim l As Long
    l = 0
    
    Dim c As String
    Do While c <> """"
        c = Mid(httpReq.responseText, InStr(httpReq.responseText, key) + Len(key) + Len(""":""") + l, 1)
        If c <> """" Then
            l = l + 1
        End If
        DoEvents
    Loop
    
    Dim newestVersion
    newestVersion = Mid(httpReq.responseText, InStr(httpReq.responseText, key) + Len(key) + Len(""":"""), l)
    
        
    If newestVersion <> thisVerion Then
        Debug.Print "最新バージョンがあります"
    End If

End Sub
