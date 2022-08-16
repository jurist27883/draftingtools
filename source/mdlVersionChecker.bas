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
        If MsgBox("�ŐV�o�[�W����������܂��B" + vbCrLf + "�_�E�����[�h���܂����B", vbYesNo) = vbYes Then
            Dim wsh As Object
            Set wsh = CreateObject("WScript.Shell")
            Dim downloadPath As String
            downloadPath = wsh.SpecialFolders("Desktop") + "\" + downloadFileName
            
            Dim strURL As String
            strURL = "https://github.com/sakura-editor/sakura/releases/download/v2.4.1/sakura-tag-v2.4.1-build2849-ee8234f-Win32-Release-Exe.zip"
            
            If URLDownloadToFile(0, strURL, downloadPath, 0, 0) = 0 Then
                MsgBox "�f�X�N�g�b�v�Ƀ_�E�����[�h���������܂����B"
            Else
                MsgBox "�t�@�C�����_�E�����[�h�ł��܂���ł����B"
            End If
        End If
    Else
        MsgBox "�ŐV�o�[�W�������C���X�g�[������Ă��܂��B"
    End If

End Sub



