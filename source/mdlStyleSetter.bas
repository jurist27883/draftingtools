Attribute VB_Name = "mdlStyleSetter"
'@Folder("main")
Option Explicit

Sub SetStyle()
    Dim p As Paragraph
    Dim styleName As String
    
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord
    
    For Each p In Selection.Paragraphs
        Select Case p.Style
        Case TITLE1, TITLE2, TITLE3, TITLE4, TITLE5, BODY1, BODY2, BODY3, BODY4, BODY5
        Case Else
            styleName = GetStyle(p)
            Select Case styleName
            Case TITLE1, TITLE2, TITLE3, TITLE4, TITLE5
                If p.Range.ListFormat.ListString = "" Then
                    p.Style = styleName
                End If
            Case Else
                Select Case GetPreiousStyle(p)
                Case TITLE1, BODY1
                    styleName = BODY1
                Case TITLE2, BODY2
                    styleName = BODY2
                Case TITLE3, BODY3
                    styleName = BODY3
                Case TITLE4, BODY4
                    styleName = BODY4
                Case TITLE5, BODY5
                    styleName = BODY5
                End Select
                p.Style = styleName
            End Select
        End Select
    Next
    
    ur.EndCustomRecord

End Sub

Private Function GetPreiousStyle(p As Paragraph) As String
    If p.Previous Is Nothing Then
        GetPreiousStyle = wdStyleNormal
        Exit Function
    End If
    
    Dim preStyle As String
    preStyle = GetStyle(p.Previous)
    
    Select Case preStyle
    Case TITLE1, TITLE2, TITLE3, TITLE4, TITLE5, BODY1, BODY2, BODY3, BODY4, BODY5
        GetPreiousStyle = preStyle
    Case Else
        GetPreiousStyle = GetPreiousStyle(p.Previous)
    End Select
End Function

Private Function GetStyle(p As Paragraph) As String
    '段落のスタイル取得・推定
    Dim char1, char2 As String
    If p.Range.ListFormat.ListString = "" Then
        char1 = Left(p, 1)
        char2 = Mid(p, 2, 1)
    Else
        char1 = Left(p.Range.ListFormat.ListString, 1)
        char2 = Mid(p.Range.ListFormat.ListString, 2, 1)
    End If
    
    Select Case p.Style
    Case TITLE1, TITLE2, TITLE3, TITLE4, TITLE5, BODY1, BODY2, BODY3, BODY4, BODY5
        GetStyle = p.Style
        Exit Function
    End Select
    
    Select Case True
    Case isOrdinalNumber(char1, char2)
        GetStyle = TITLE1
    Case isIndexNumber(char1, char2)
        GetStyle = TITLE2
    Case isBracketsNumber(char1, char2)
        GetStyle = TITLE3
    Case isIndexKatakana(char1, char2)
        GetStyle = TITLE4
    Case isBracketsKatakana(char1, char2)
        GetStyle = TITLE5
    Case Else
        GetStyle = p.Style
    End Select
        
End Function

Private Function isBracketsKatakana(ByVal s1 As String, ByVal s2 As String) As Boolean
    Select Case s1
    Case "(", "（"
        If isKatakana(s2) Then
            isBracketsKatakana = True
        Else
            isBracketsKatakana = False
        End If
    Case Else
        isBracketsKatakana = False
    End Select
End Function

Private Function isIndexKatakana(ByVal s1 As String, ByVal s2 As String) As Boolean
    If isKatakana(s1) And isSpace(s2) Then
        isIndexKatakana = True
    Else
        isIndexKatakana = False
    End If
End Function

Private Function isIndexNumber(ByVal s1 As String, ByVal s2 As String) As Boolean
    If IsNumeric(s1) And isSpace(s2) Then
        isIndexNumber = True
    Else
        isIndexNumber = False
    End If
End Function

Private Function isOrdinalNumber(ByVal s1 As String, ByVal s2 As String) As Boolean
    '第
    If s1 = "第" Then
        If IsNumeric(s2) Then
            isOrdinalNumber = True
        Else
            isOrdinalNumber = False
        End If
    Else
        isOrdinalNumber = False
    End If
End Function

Private Function isBracketsNumber(ByVal s1 As String, ByVal s2 As String) As Boolean
    '括弧数字一文字
    Select Case AscW(s1)
    Case -8191 To -8093, 9332 To 9351
        isBracketsNumber = True
    Case Else
        Select Case s1
        Case "(", "（"
            If IsNumeric(s2) Then
                isBracketsNumber = True
            Else
                isBracketsNumber = False
            End If
        Case Else
            isBracketsNumber = False
        End Select
    End Select
End Function


