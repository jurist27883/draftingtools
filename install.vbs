On Error Resume Next

Dim installPath
Dim addinName
Dim addinFile

addinName = "�N��"
addinFile = "draftingtools.dotm" 

IF MsgBox(addinName & "�A�h�C�����C���X�g�[�����܂����H", vbYesNo + vbQuestion) = vbNo Then
  WScript.Quit
End IF

Set wshShell = CreateObject("WScript.Shell")
Set fileSystem = CreateObject("Scripting.FileSystemObject")

installPath = wshShell.SpecialFolders("Appdata") & "\Microsoft\Word\Startup\" & addinFile

fileSystem.CopyFile  addinFile ,installPath , True

Set wshShell = Nothing
Set fileSystem = Nothing

IF Err.Number = 0 THEN
   MsgBox "�A�h�C���̃C���X�g�[�����������܂����B", vbInformation
ELSE
   MsgBox "�G���[���������܂����Bword����čĎ��s���Ă��������B" , vbExclamation
End IF