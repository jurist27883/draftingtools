On Error Resume Next

Dim installPath
Dim addinName
Dim addinFile

addinName = "起案"
addinFile = "draftingtools.dotm" 

IF MsgBox(addinName & "アドインをインストールしますか？", vbYesNo + vbQuestion) = vbNo Then
  WScript.Quit
End IF

Set wshShell = CreateObject("WScript.Shell")
Set fileSystem = CreateObject("Scripting.FileSystemObject")

installPath = wshShell.SpecialFolders("Appdata") & "\Microsoft\Word\Startup\" & addinFile

fileSystem.CopyFile  addinFile ,installPath , True

Set wshShell = Nothing
Set fileSystem = Nothing

IF Err.Number = 0 THEN
   MsgBox "アドインのインストールが完了しました。", vbInformation
ELSE
   MsgBox "エラーが発生しました。wordを閉じて再実行してください。" , vbExclamation
End IF