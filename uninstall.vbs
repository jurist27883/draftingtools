On Error Resume Next

Dim installPath
Dim addinName
Dim addinFile

addinName = "起案"
addinFile = "draftingtools.dotm" 

If MsgBox(addinName & "アドインをアンインストールしますか？", vbYesNo + vbQuestion) = vbNo Then
  WScript.Quit
End If

Set wshShell = CreateObject("WScript.Shell")
Set fileSystem = CreateObject("Scripting.FileSystemObject")

installPath = wshShell.SpecialFolders("Appdata") & "\Microsoft\Word\Startup\" & addinFile

If fileSystem.FileExists(installPath) = True Then
  fileSystem.DeleteFile installPath , True
Else
  MsgBox "既にアンインストール済みです。", vbExclamation
End If

Set wshShell = Nothing
Set fileSystem = Nothing

If Err.Number = 0 Then
   MsgBox "アドインのアンインストールが完了しました。", vbInformation
Else
   MsgBox "エラーが発生しました。wordを閉じてから再度実行してください。", vbExclamation
End If
