On Error Resume Next

Dim installPath
Dim addinName
Dim addinFile

addinName = "�N��"
addinFile = "draftingtools.dotm" 

If MsgBox(addinName & "�A�h�C�����A���C���X�g�[�����܂����H", vbYesNo + vbQuestion) = vbNo Then
  WScript.Quit
End If

Set wshShell = CreateObject("WScript.Shell")
Set fileSystem = CreateObject("Scripting.FileSystemObject")

installPath = wshShell.SpecialFolders("Appdata") & "\Microsoft\Word\Startup\" & addinFile

If fileSystem.FileExists(installPath) = True Then
  fileSystem.DeleteFile installPath , True
Else
  MsgBox "���ɃA���C���X�g�[���ς݂ł��B", vbExclamation
End If

Set wshShell = Nothing
Set fileSystem = Nothing

If Err.Number = 0 Then
   MsgBox "�A�h�C���̃A���C���X�g�[�����������܂����B", vbInformation
Else
   MsgBox "�G���[���������܂����Bword����Ă���ēx���s���Ă��������B", vbExclamation
End If
