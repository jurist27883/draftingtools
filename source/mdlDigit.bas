Attribute VB_Name = "mdlDigit"
'@Folder "main"
Option Explicit

Sub ChangeDigit()
    Dim targetRange As Range
    Set targetRange = ExpandRange(Selection.Range, "0123456789‚O‚P‚Q‚R‚S‚T‚U‚V‚W‚X‚O’›‰­–œ,C")
    
    If targetRange.text = "" Then
        Exit Sub
    End If
    
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord
    
    If InStr(targetRange.text, "’›") <> 0 Or InStr(targetRange.text, "‰­") _
        Or InStr(targetRange.text, "–œ") Then
        OkumanToComma targetRange
    Else
        CommaToOkuman targetRange
    End If
        
    ur.EndCustomRecord
End Sub

Sub ArabicToChinese()
    'ƒAƒ‰ƒrƒA”š‚ğŠ¿”š‚É
    
    '‘I‘ğ”ÍˆÍ‚ğŠgk‚·‚é
    Dim targetRange As Range
    Set targetRange = ExpandRange(Selection.Range, "0123456789‚O‚P‚Q‚R‚S‚T‚U‚V‚W‚X‚O,C")
    
    '‘I‘ğ”ÍˆÍ‚Ì•¶š—ñ‚ª‹ó‚È‚çI—¹
    If targetRange.text = "" Then
        Exit Sub
    End If
    
    With targetRange
        Dim sourceText As String
        sourceText = .text
        
        With .Find
            .text = sourceText
            .Replacement.text = ConvertToChinese(sourceText)
            .Execute Replace:=wdReplaceAll
        End With
    End With
End Sub

Private Function ConvertToChinese(sourceText As String) As String
    'Š¿”š•ÏŠ·
    'ç•S‚Ì•t‰Á
    sourceText = Format(sourceText, "#ç#•S#\0")
    Dim length
    length = Len(sourceText)
    
    Dim i As Long
    For i = 1 To length
        If Left(sourceText, 1) Like "[ç,•S,\]" Then
            sourceText = Right(sourceText, length - i)
        Else
            Exit For
        End If
    Next
    
    '–`“ªu1v‚Ìœ‹
    If Len(sourceText) > 1 And Left(sourceText, 1) = "1" Then
        sourceText = Right(sourceText, Len(sourceText) - 1)
    End If
    
    '’PŠ¿”š•ÏŠ·
    For i = 1 To 9
        sourceText = Replace(sourceText, CStr(i), Mid("ˆê“ñOlŒÜ˜Zµ”ª‹ã", i, 1))
    Next
    
     '“r’†‚Ì0‚ÆŸ‚Ì‹æØ‚è‚Ìœ‹
    If Right(sourceText, 1) = "0" Then
        sourceText = Left(sourceText, Len(sourceText) - 1)
    End If
    If Right(sourceText, 2) = "0\" Then
        sourceText = Left(sourceText, Len(sourceText) - 2)
    End If
    If Right(sourceText, 2) = "0•S" Then
        sourceText = Left(sourceText, Len(sourceText) - 2)
    End If
     
    ConvertToChinese = sourceText
    
End Function

Sub ChineseToArabic()
    'Š¿”š‚ğƒAƒ‰ƒrƒA”š‚É
    
    '‘I‘ğ”ÍˆÍ‚ğŠgk‚·‚é
    Dim targetRange As Range
    Set targetRange = ExpandRange(Selection.Range, "ˆê“ñOlŒÜ˜Zµ”ª‹ã\•Sç")
    
    '’uŠ·
    With targetRange
        Dim sourceText As String
        sourceText = .text
        With .Find
            .text = sourceText
            .Replacement.text = Format(ConvertToArabic(sourceText), "###0")
            .Execute Replace:=wdReplaceAll
        End With
        .CharacterWidth = wdWidthFullWidth
    End With
    
End Sub

Private Function ConvertToArabic(sourceText As String) As String
    'ƒAƒ‰ƒrƒA”š•ÏŠ·
    
    Dim i, j As Long
    For i = 1 To 9
        sourceText = Replace(sourceText, Mid("ˆê“ñOlŒÜ˜Zµ”ª‹ã", i, 1), CStr(i))
    Next
    
    sourceText = AddInitial(sourceText)

    Dim splited() As String
    Dim delimiter As String
    Dim digits As String
    Dim digitNumbers(3) As String
    digits = "ç•S\"
    
    For i = 0 To 2
        delimiter = Mid(digits, i + 1, 1)
        splited = Split(sourceText, delimiter)
        
        Select Case UBound(splited)
        Case -1
            For j = i To 3
                digitNumbers(j) = "0"
            Next
        Case 0
            digitNumbers(i) = "0"
            digitNumbers(i + 1) = splited(0)
            sourceText = AddInitial(splited(0))
        Case Else
            digitNumbers(i) = splited(0)
            
            If splited(1) = "" Then
                splited(1) = 0
            End If
            digitNumbers(i + 1) = splited(1)
            sourceText = AddInitial(splited(1))
        End Select
    Next
    
    For i = 0 To 3
        ConvertToArabic = ConvertToArabic + digitNumbers(i)
    Next

End Function
Private Function AddInitial(sourceText As String) As String
    If Left(sourceText, 1) = "ç" Or Left(sourceText, 1) = "•S" Or Left(sourceText, 1) = "\" Then
        AddInitial = "1" + sourceText
    Else
        AddInitial = sourceText
    End If
End Function

Private Sub OkumanToComma(targetRange As Range)
    '‰­–œ‹æØ‚è‚ğƒRƒ“ƒ}‹æØ‚è‚É
    
    '‘I‘ğ”ÍˆÍ‚ğŠgk‚·‚é
'    Dim targetRange As Range
'    Set targetRange = ExpandRange(Selection.Range, "0123456789‚O‚P‚Q‚R‚S‚T‚U‚V‚W‚X‚O’›‰­–œ")
        
    '’uŠ·
    With targetRange
        Dim sourceText As String
        sourceText = .text
        
        With .Find
            .text = sourceText
            .Replacement.text = Format(ConvertToComma(sourceText), "#,##0")
            .Execute Replace:=wdReplaceAll
        End With
    End With
    
End Sub

Private Function ConvertToComma(sourceText As String) As String
    '‰­–œ‹æØ‚è‚©‚çƒRƒ“ƒ}‹æØ‚è
    
    Dim splited() As String
    Dim delimiter As String
    Dim digits As String
    Dim digitNumbers(3) As String
    digits = "’›‰­–œ"
    Dim i, j As Long
    
    For i = 0 To 2
        delimiter = Mid(digits, i + 1, 1)
        splited = Split(sourceText, delimiter)
        
        Select Case UBound(splited)
        Case -1
            For j = i To 3
                digitNumbers(j) = "0000"
            Next
        Case 0
            digitNumbers(i) = "0000"
            digitNumbers(i + 1) = Format(splited(0), "0000")
            sourceText = splited(0)
        Case Else
            digitNumbers(i) = Format(splited(0), "0000")
            
            If splited(1) = "" Then
                splited(1) = 0
            End If
            digitNumbers(i + 1) = Format(splited(1), "0000")
            sourceText = splited(1)
        End Select
    Next
    
    For i = 0 To 3
        ConvertToComma = ConvertToComma + digitNumbers(i)
    Next
    
End Function

Private Sub CommaToOkuman(targetRange As Range)
    'ƒRƒ“ƒ}‹æØ‚è‚ğ‰­–œ‹æØ‚è‚É
    
    '‘I‘ğ”ÍˆÍ‚ğŠgk‚·‚é
'    Dim targetRange As Range
'    Set targetRange = ExpandRange(Selection.Range, "0123456789‚O‚P‚Q‚R‚S‚T‚U‚V‚W‚X‚O,C")
    
    '‘I‘ğ”ÍˆÍ‚Ì•¶š—ñ‚ª‹ó‚È‚çI—¹
    If targetRange.text = "" Then
        Exit Sub
    End If
    
    With targetRange
        Dim sourceText As String
        sourceText = .text
        
        '’uŠ·
        With .Find
            .text = sourceText
            .Replacement.text = ConvertToOkuman(sourceText)
            .Execute Replace:=wdReplaceAll
        End With
    End With
End Sub

Private Function ConvertToOkuman(sourceText As String) As String
        
    'ƒRƒ“ƒ}‹æØ‚è‚©‚ç‰­–œ‹æØ‚è
    Dim TargetChar As String
    
    '‰­–œ‚Ì•t‰Á
    sourceText = Format(sourceText, "####’›####‰­####–œ###0")
    
    Dim length
    length = Len(sourceText)
    
    Dim i As Long
    For i = 1 To length
        If Left(sourceText, 1) Like "[’›,‰­,–œ]" Then
            sourceText = Right(sourceText, length - i)
        Else
            Exit For
        End If
    Next
    
    '“r’†‚Ì0‚ÆŸ‚Ì‹æØ‚è‚Ìœ‹
    If Right(sourceText, 4) = "0000" Then
        sourceText = Left(sourceText, Len(sourceText) - 4)
    End If
     
    If Right(sourceText, 5) = "0000–œ" Then
        sourceText = Left(sourceText, Len(sourceText) - 5)
    End If
    If Right(sourceText, 5) = "0000‰­" Then
        sourceText = Left(sourceText, Len(sourceText) - 5)
    End If
     
    '‘SŠp‚É•ÏŠ·
    ConvertToOkuman = StrConv(sourceText, vbWide)
     
End Function

Private Function ExpandRange(targetRange As Range, charSet As String) As Range
    '‘I‘ğ”ÍˆÍ‚ÌŠgk
    With targetRange
        '‘I‘ğ”ÍˆÍ‚ğL‚°‚é
        .MoveStartWhile charSet, wdBackward
        .MoveEndWhile charSet
        '‘I‘ğ”ÍˆÍ‚ğ‹·‚ß‚é
        .MoveStartUntil charSet
        .MoveEndUntil charSet, wdBackward
    End With
    
    Set ExpandRange = targetRange
    
End Function
