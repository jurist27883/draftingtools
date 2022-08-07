Attribute VB_Name = "mdlRevision"
'@Folder "main"
Option Explicit

Sub ToggleTrackingFormat()
    ' �����ύX�����L�^�ؑւ�
    ActiveDocument.TrackFormatting = Not ActiveDocument.TrackFormatting
End Sub

Sub AcceptChangedFormat()
    ' �����ύX���𔽉f
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
    '���Ȃ̕ύX������S�Ĕ��f
    Dim r As Revision
    For Each r In ActiveDocument.Revisions
        If r.Author = Application.UserName Then
            r.Accept
        End If
    Next
End Sub
