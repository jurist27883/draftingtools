Attribute VB_Name = "mdlStyleSetter"
'@Folder "main.style"
Option Explicit

Sub SetStyles()
<<<<<<< HEAD
=======
    'é©•ï½¸è¬šæ¨’ï½¯?¿½è—ï½²ç¸º?½«ç¹§?½¹ç¹§?½¿ç¹§?½¤ç¹ï½«é–¾?½ªèœæ?Šï½¨?½­èž³?¿½
    Dim p As Paragraph
    Dim styleName As String
    
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord
    
    For Each p In Selection.Paragraphs
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
    'è°¿?½µé—œï½½ç¸º?½®ç¹§?½¹ç¹§?½¿ç¹§?½¤ç¹ï½«èœ¿é–€?½¾åŠ±?¿½?½»è¬—ï½¨èž³?¿½
    Dim char1, char2 As String
    If p.Range.ListFormat.ListString = "" Then
        char1 = Left(p, 1)
        char2 = Mid(p, 2, 1)
    Else
        char1 = Left(p.Range.ListFormat.ListString, 1)
        char2 = Mid(p.Range.ListFormat.ListString, 2, 1)
    End If
    
    Select Case True
    Case IsOrdinalNumber(char1, char2)
        GetStyle = TITLE1
    Case IsIndexNumber(char1, char2)
        GetStyle = TITLE2
    Case IsBracketsNumber(char1, char2)
        GetStyle = TITLE3
    Case IsIndexKatakana(char1, char2)
        GetStyle = TITLE4
    Case IsBracketsKatakana(char1, char2)
        GetStyle = TITLE5
    Case Else
        GetStyle = p.Style
    End Select
        
End Function

Private Function IsBracketsKatakana(ByVal s1 As String, ByVal s2 As String) As Boolean
    Select Case s1
    Case "(", "?¿½?½¼?¿½"
        If isKatakana(s2) Then
            IsBracketsKatakana = True
        Else
            IsBracketsKatakana = False
        End If
    Case Else
        IsBracketsKatakana = False
    End Select
End Function

Private Function IsIndexKatakana(ByVal s1 As String, ByVal s2 As String) As Boolean
    If isKatakana(s1) And isSpace(s2) Then
        IsIndexKatakana = True
    Else
        IsIndexKatakana = False
    End If
End Function

Private Function IsIndexNumber(ByVal s1 As String, ByVal s2 As String) As Boolean
    If IsNumeric(s1) And isSpace(s2) Then
        IsIndexNumber = True
    Else
        IsIndexNumber = False
    End If
End Function

Private Function IsOrdinalNumber(ByVal s1 As String, ByVal s2 As String) As Boolean
    If s1 = "éš¨?½¬" Then
        If IsNumeric(s2) Then
            IsOrdinalNumber = True
        Else
            IsOrdinalNumber = False
        End If
    Else
        IsOrdinalNumber = False
    End If
End Function

Private Function IsBracketsNumber(ByVal s1 As String, ByVal s2 As String) As Boolean
    'è«¡?½¬è ‘ï½§è¬¨?½°èŸ?å¶º?½¸Â€è­??¿½èŸ??¿½
    Select Case AscW(s1)
    Case -8191 To -8093, 9332 To 9351
        IsBracketsNumber = True
    Case Else
        Select Case s1
        Case "(", "?¿½?½¼?¿½"
            If IsNumeric(s2) Then
                IsBracketsNumber = True
            Else
                IsBracketsNumber = False
            End If
        Case Else
            IsBracketsNumber = False
        End Select
    End Select
End Function

