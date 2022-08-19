Attribute VB_Name = "mdlVersionChecker"
'@Folder "version"
Option Explicit

Const thisVerion = "v1.0.0"
Const key = "tag_name"
Const checkUrl = "https://api.github.com/repos/jurist27883/draftingtools/releases/latest"
Const downloadUrl = "https://github.com/jurist27883/draftingtools/releases/tag/"
Const downloadFileName = "draftingtools.zip"

#If VBA7 Then
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
   (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#Else
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
   (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

Sub CheckUpdate()
    
    #If RELEASE Then
        Dim httpReq As Object
        Set httpReq = CreateObject("MSXML2.ServerXMLHTTP")
    #Else
        Dim httpReq As MSXML2.XMLHTTP60
        Set httpReq = New MSXML2.XMLHTTP60
    #End If

    With httpReq
        .Open "GET", checkUrl
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
<<<<<<< HEAD
        If MsgBox("ï¿½ÅVï¿½oï¿½[ï¿½Wï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Ü‚ï¿½ï¿½B" + vbCrLf + "ï¿½_ï¿½Eï¿½ï¿½ï¿½ï¿½ï¿½[ï¿½hï¿½ï¿½ï¿½Ü‚ï¿½ï¿½ï¿½ï¿½B", vbYesNo) = vbYes Then
=======
        If MsgBox("ÅVƒo[ƒWƒ‡ƒ“‚ª‚ ‚è‚Ü‚·B" + vbCrLf + "ƒ_ƒEƒ“ƒ[ƒh‚µ‚Ü‚·‚©B", vbYesNo) = vbYes Then
>>>>>>> 514445e (release)
            Dim wsh As Object
            Set wsh = CreateObject("WScript.Shell")
            Dim downloadPath As String
            downloadPath = wsh.SpecialFolders("Desktop") + "\" + downloadFileName
            
            Dim strURL As String
            strURL = "https://github.com/sakura-editor/sakura/releases/download/v2.4.1/sakura-tag-v2.4.1-build2849-ee8234f-Win32-Release-Exe.zip"
            
            If URLDownloadToFile(0, strURL, downloadPath, 0, 0) = 0 Then
                MsgBox "ï¿½fï¿½Xï¿½Nï¿½gï¿½bï¿½vï¿½Éƒ_ï¿½Eï¿½ï¿½ï¿½ï¿½ï¿½[ï¿½hï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Ü‚ï¿½ï¿½ï¿½ï¿½B"
            Else
                MsgBox "ï¿½tï¿½@ï¿½Cï¿½ï¿½ï¿½ï¿½ï¿½_ï¿½Eï¿½ï¿½ï¿½ï¿½ï¿½[ï¿½hï¿½Å‚ï¿½ï¿½Ü‚ï¿½ï¿½ï¿½Å‚ï¿½ï¿½ï¿½ï¿½B"
            End If
        End If
    Else
<<<<<<< HEAD
        MsgBox "ï¿½ÅVï¿½oï¿½[ï¿½Wï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Cï¿½ï¿½ï¿½Xï¿½gï¿½[ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Ä‚ï¿½ï¿½Ü‚ï¿½ï¿½B"
=======
        MsgBox "ÅVƒo[ƒWƒ‡ƒ“‚ªƒCƒ“ƒXƒg[ƒ‹‚³‚ê‚Ä‚¢‚Ü‚·B"
>>>>>>> 514445e (release)
    End If

End Sub



