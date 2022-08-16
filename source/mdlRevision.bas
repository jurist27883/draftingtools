Attribute VB_Name = "mdlRevision"
'@Folder "main"
Option Explicit

Sub ToggleTrackingFormat()
    ' ‘®•ÏX—š—ğ‹L˜^Ø‘Ö‚¦
    ActiveDocument.TrackFormatting = Not ActiveDocument.TrackFormatting
End Sub

Sub AcceptChangedFormat()
    ' ‘®•ÏX—š—ğ”½‰f
    Dim r As Revision
    For Each r In ActiveDocument.Revisions
        Select Case r.Type
        Case wdRevisionParagraphProperty, wdRevisionSectionProperty, wdRevisionTableProperty, wdRevisionProperty
            r.Accept
        End Select
    Next
    ActiveDocument.TrackFormatting = False
End Sub

Sub AcceptRevisions()
    '•ÏX—š—ğ‚ğ”½‰f
    Dim editors As Variant
    editors = GetEditor
    
    Dim i, j As Long
    Dim editorName As String
    ReDim settledEditors(0) As Variant
    Dim r As Revision
    
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord
    
    For i = 0 To UBound(editors)
        editorName = editors(i)
        For j = 0 To UBound(settledEditors)
            If editorName = settledEditors(j) Then
                GoTo Continue
            End If
        Next
        If MsgBox(editorName + "‚Ì•ÏX—š—ğ‚ğ”½‰f‚³‚¹‚Ü‚·‚©", vbYesNo) = vbYes Then
            For Each r In ActiveDocument.Revisions
                If r.Author = editorName Then
                    r.Accept
                End If
            Next
        End If
        PushArray settledEditors, editorName
Continue:
    Next
    ur.EndCustomRecord
End Sub

Private Function GetEditor() As Variant
    Dim r As Revision
    ReDim editors(0) As Variant
    For Each r In Selection.Range.Revisions
        PushArray editors, r.Author
    Next
    GetEditor = editors
End Function
