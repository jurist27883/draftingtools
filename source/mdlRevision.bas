Attribute VB_Name = "mdlRevision"
'@Folder "main"
Option Explicit

Sub ToggleTrackingFormat()
    ' 書式変更履歴記録切替え
    ActiveDocument.TrackFormatting = Not ActiveDocument.TrackFormatting
End Sub

Sub AcceptChangedFormat()
    ' 書式変更履歴反映
    Dim r As Revision
    For Each r In ActiveDocument.Revisions
        Select Case r.Type
        Case wdRevisionParagraphProperty, wdRevisionSectionProperty, wdRevisionTableProperty, wdRevisionProperty
            r.Accept
        End Select
    Next
    ActiveDocument.TrackFormatting = False
End Sub


Sub AcceptMyRevision()
    '自己の変更履歴を全て反映
    Dim r As Revision
    For Each r In ActiveDocument.Revisions
        If r.Author = Application.UserName Then
            r.Accept
        End If
    Next
End Sub
