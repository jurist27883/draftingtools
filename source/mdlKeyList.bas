Attribute VB_Name = "mdlKeyList"
'@Folder("config")
Option Explicit

Sub InitializeKeyListBox()
    With frmConfig.lstKeys

        Dim i As Long
        For i = 0 To 8
            .AddItem ""
            .List(i, 0) = i
            .List(i, 1) = i + 1
        Next
    End With
End Sub
