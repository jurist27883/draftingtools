Attribute VB_Name = "mdlDigit"
'@Folder "main"
Option Explicit
Sub ArabicToChinese()
    'ƒAƒ‰ƒrƒA”š‚ğŠ¿”š‚É
    
    '‘I‘ğ”ÍˆÍ‚ğŠgk‚·‚é
    Dim TargetRange As Range
    Set TargetRange = ExpandRange(Selection.Range, "0123456789‚O‚P‚Q‚R‚S‚T‚U‚V‚W‚X‚O,C")
    
    '‘I‘ğ”ÍˆÍ‚Ì•¶š—ñ‚ª‹ó‚È‚çI—¹
    If TargetRange.text = "" Then
        Exit Sub
    End If
    
    With TargetRange
        Dim SourceText As String
        SourceText = .text
        
        With .Find
            .text = SourceText
            .Replacement.text = ConvertToChinese(SourceText)
            .Execute Replace:=wdReplaceAll
        End With
    End With
End Sub

Private Function ConvertToChinese(SourceText As String) As String
    'Š¿”š•ÏŠ·
    'ç•S‚Ì•t‰Á
    SourceText = Format(SourceText, "#ç#•S#\0")
    Dim Length
    Length = Len(SourceText)
    
    Dim i As Long
    For i = 1 To Length
        If Left(SourceText, 1) Like "[ç,•S,\]" Then
            SourceText = Right(SourceText, Length - i)
        Else
            Exit For
        End If
    Next
    
    '–`“ªu1v‚Ìœ‹
    If Len(SourceText) > 1 And Left(SourceText, 1) = "1" Then
        SourceText = Right(SourceText, Len(SourceText) - 1)
    End If
    
    '’PŠ¿”š•ÏŠ·
    For i = 1 To 9
        SourceText = Replace(SourceText, CStr(i), Mid("ˆê“ñOlŒÜ˜Zµ”ª‹ã", i, 1))
    Next
    
     '“r’†‚Ì0‚ÆŸ‚Ì‹æØ‚è‚Ìœ‹
    If Right(SourceText, 1) = "0" Then
        SourceText = Left(SourceText, Len(SourceText) - 1)
    End If
    If Right(SourceText, 2) = "0\" Then
        SourceText = Left(SourceText, Len(SourceText) - 2)
    End If
    If Right(SourceText, 2) = "0•S" Then
        SourceText = Left(SourceText, Len(SourceText) - 2)
    End If
     
    ConvertToChinese = SourceText
    
End Function

Sub ChineseToArabic()
    'Š¿”š‚ğƒAƒ‰ƒrƒA”š‚É
    
    '‘I‘ğ”ÍˆÍ‚ğŠgk‚·‚é
    Dim TargetRange As Range
    Set TargetRange = ExpandRange(Selection.Range, "ˆê“ñOlŒÜ˜Zµ”ª‹ã\•Sç")
    
    '’uŠ·
    With TargetRange
        Dim SourceText As String
        SourceText = .text
        With .Find
            .text = SourceText
            .Replacement.text = Format(ConvertToArabic(SourceText), "###0")
            .Execute Replace:=wdReplaceAll
        End With
        .CharacterWidth = wdWidthFullWidth
    End With
    
End Sub

Private Function ConvertToArabic(SourceText As String) As String
    'ƒAƒ‰ƒrƒA”š•ÏŠ·
    
    Dim i, j As Long
    For i = 1 To 9
        SourceText = Replace(SourceText, Mid("ˆê“ñOlŒÜ˜Zµ”ª‹ã", i, 1), CStr(i))
    Next
    
    SourceText = AddInitial(SourceText)

    Dim Splited() As String
    Dim Delimiter As String
    Dim Digits As String
    Dim DigitNumbers(3) As String
    Digits = "ç•S\"
    
    For i = 0 To 2
        Delimiter = Mid(Digits, i + 1, 1)
        Splited = Split(SourceText, Delimiter)
        
        Select Case UBound(Splited)
        Case -1
            For j = i To 3
                DigitNumbers(j) = "0"
            Next
        Case 0
            DigitNumbers(i) = "0"
            DigitNumbers(i + 1) = Splited(0)
            SourceText = AddInitial(Splited(0))
        Case Else
            DigitNumbers(i) = Splited(0)
            
            If Splited(1) = "" Then
                Splited(1) = 0
            End If
            DigitNumbers(i + 1) = Splited(1)
            SourceText = AddInitial(Splited(1))
        End Select
    Next
    
    For i = 0 To 3
        ConvertToArabic = ConvertToArabic + DigitNumbers(i)
    Next

End Function
Private Function AddInitial(SourceText As String) As String
    If Left(SourceText, 1) = "ç" Or Left(SourceText, 1) = "•S" Or Left(SourceText, 1) = "\" Then
        AddInitial = "1" + SourceText
    Else
        AddInitial = SourceText
    End If
End Function

Sub OkumanToComma()
    '‰­–œ‹æØ‚è‚ğƒRƒ“ƒ}‹æØ‚è‚É
    
    '‘I‘ğ”ÍˆÍ‚ğŠgk‚·‚é
    Dim TargetRange As Range
    Set TargetRange = ExpandRange(Selection.Range, "0123456789‚O‚P‚Q‚R‚S‚T‚U‚V‚W‚X‚O’›‰­–œ")
        
    '’uŠ·
    With TargetRange
        Dim SourceText As String
        SourceText = .text
        
        With .Find
            .text = SourceText
            .Replacement.text = Format(ConvertToComma(SourceText), "#,##0")
            .Execute Replace:=wdReplaceAll
        End With
    End With
    
End Sub

Private Function ConvertToComma(SourceText As String) As String
    '‰­–œ‹æØ‚è‚©‚çƒRƒ“ƒ}‹æØ‚è
    
    Dim Splited() As String
    Dim Delimiter As String
    Dim Digits As String
    Dim DigitNumbers(3) As String
    Digits = "’›‰­–œ"
    Dim i, j As Long
    
    For i = 0 To 2
        Delimiter = Mid(Digits, i + 1, 1)
        Splited = Split(SourceText, Delimiter)
        
        Select Case UBound(Splited)
        Case -1
            For j = i To 3
                DigitNumbers(j) = "0000"
            Next
        Case 0
            DigitNumbers(i) = "0000"
            DigitNumbers(i + 1) = Format(Splited(0), "0000")
            SourceText = Splited(0)
        Case Else
            DigitNumbers(i) = Format(Splited(0), "0000")
            
            If Splited(1) = "" Then
                Splited(1) = 0
            End If
            DigitNumbers(i + 1) = Format(Splited(1), "0000")
            SourceText = Splited(1)
        End Select
    Next
    
    For i = 0 To 3
        ConvertToComma = ConvertToComma + DigitNumbers(i)
    Next
    
End Function

Sub CommaToOkuman()
    'ƒRƒ“ƒ}‹æØ‚è‚ğ‰­–œ‹æØ‚è‚É
    
    '‘I‘ğ”ÍˆÍ‚ğŠgk‚·‚é
    Dim TargetRange As Range
    Set TargetRange = ExpandRange(Selection.Range, "0123456789‚O‚P‚Q‚R‚S‚T‚U‚V‚W‚X‚O,C")
    
    '‘I‘ğ”ÍˆÍ‚Ì•¶š—ñ‚ª‹ó‚È‚çI—¹
    If TargetRange.text = "" Then
        Exit Sub
    End If
    
    With TargetRange
        Dim SourceText As String
        SourceText = .text
        
        '’uŠ·
        With .Find
            .text = SourceText
            .Replacement.text = ConvertToOkuman(SourceText)
            .Execute Replace:=wdReplaceAll
        End With
    End With
End Sub

Private Function ConvertToOkuman(SourceText As String) As String
        
    'ƒRƒ“ƒ}‹æØ‚è‚©‚ç‰­–œ‹æØ‚è
    Dim TargetChar As String
    
    '‰­–œ‚Ì•t‰Á
    SourceText = Format(SourceText, "####’›####‰­####–œ###0")
    
    Dim Length
    Length = Len(SourceText)
    
    Dim i As Long
    For i = 1 To Length
        If Left(SourceText, 1) Like "[’›,‰­,–œ]" Then
            SourceText = Right(SourceText, Length - i)
        Else
            Exit For
        End If
    Next
    
    '“r’†‚Ì0‚ÆŸ‚Ì‹æØ‚è‚Ìœ‹
    If Right(SourceText, 4) = "0000" Then
        SourceText = Left(SourceText, Len(SourceText) - 4)
    End If
     
    If Right(SourceText, 5) = "0000–œ" Then
        SourceText = Left(SourceText, Len(SourceText) - 5)
    End If
    If Right(SourceText, 5) = "0000‰­" Then
        SourceText = Left(SourceText, Len(SourceText) - 5)
    End If
     
    '‘SŠp‚É•ÏŠ·
    ConvertToOkuman = StrConv(SourceText, vbWide)
     
End Function

Private Function ExpandRange(TargetRange As Range, CharSet As String) As Range
    '‘I‘ğ”ÍˆÍ‚ÌŠgk
    With TargetRange
        '‘I‘ğ”ÍˆÍ‚ğL‚°‚é
        .MoveStartWhile CharSet, wdBackward
        .MoveEndWhile CharSet
        '‘I‘ğ”ÍˆÍ‚ğ‹·‚ß‚é
        .MoveStartUntil CharSet
        .MoveEndUntil CharSet, wdBackward
    End With
    
    Set ExpandRange = TargetRange
    
End Function
