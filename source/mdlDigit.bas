Attribute VB_Name = "mdlDigit"
'@Folder "main"
Option Explicit
Sub ArabicToChinese()
    'アラビア数字を漢数字に
    
    '選択範囲を拡縮する
    Dim TargetRange As Range
    Set TargetRange = ExpandRange(Selection.Range, "0123456789０１２３４５６７８９０,，")
    
    '選択範囲の文字列が空なら終了
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
    '漢数字変換
    '千百の付加
    SourceText = Format(SourceText, "#千#百#十0")
    Dim Length
    Length = Len(SourceText)
    
    Dim i As Long
    For i = 1 To Length
        If Left(SourceText, 1) Like "[千,百,十]" Then
            SourceText = Right(SourceText, Length - i)
        Else
            Exit For
        End If
    Next
    
    '冒頭「1」の除去
    If Len(SourceText) > 1 And Left(SourceText, 1) = "1" Then
        SourceText = Right(SourceText, Len(SourceText) - 1)
    End If
    
    '単漢数字変換
    For i = 1 To 9
        SourceText = Replace(SourceText, CStr(i), Mid("一二三四五六七八九", i, 1))
    Next
    
     '途中の0と次の区切りの除去
    If Right(SourceText, 1) = "0" Then
        SourceText = Left(SourceText, Len(SourceText) - 1)
    End If
    If Right(SourceText, 2) = "0十" Then
        SourceText = Left(SourceText, Len(SourceText) - 2)
    End If
    If Right(SourceText, 2) = "0百" Then
        SourceText = Left(SourceText, Len(SourceText) - 2)
    End If
     
    ConvertToChinese = SourceText
    
End Function

Sub ChineseToArabic()
    '漢数字をアラビア数字に
    
    '選択範囲を拡縮する
    Dim TargetRange As Range
    Set TargetRange = ExpandRange(Selection.Range, "一二三四五六七八九十百千")
    
    '置換
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
    'アラビア数字変換
    
    Dim i, j As Long
    For i = 1 To 9
        SourceText = Replace(SourceText, Mid("一二三四五六七八九", i, 1), CStr(i))
    Next
    
    SourceText = AddInitial(SourceText)

    Dim Splited() As String
    Dim Delimiter As String
    Dim Digits As String
    Dim DigitNumbers(3) As String
    Digits = "千百十"
    
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
    If Left(SourceText, 1) = "千" Or Left(SourceText, 1) = "百" Or Left(SourceText, 1) = "十" Then
        AddInitial = "1" + SourceText
    Else
        AddInitial = SourceText
    End If
End Function

Sub OkumanToComma()
    '億万区切りをコンマ区切りに
    
    '選択範囲を拡縮する
    Dim TargetRange As Range
    Set TargetRange = ExpandRange(Selection.Range, "0123456789０１２３４５６７８９０兆億万")
        
    '置換
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
    '億万区切りからコンマ区切り
    
    Dim Splited() As String
    Dim Delimiter As String
    Dim Digits As String
    Dim DigitNumbers(3) As String
    Digits = "兆億万"
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
    'コンマ区切りを億万区切りに
    
    '選択範囲を拡縮する
    Dim TargetRange As Range
    Set TargetRange = ExpandRange(Selection.Range, "0123456789０１２３４５６７８９０,，")
    
    '選択範囲の文字列が空なら終了
    If TargetRange.text = "" Then
        Exit Sub
    End If
    
    With TargetRange
        Dim SourceText As String
        SourceText = .text
        
        '置換
        With .Find
            .text = SourceText
            .Replacement.text = ConvertToOkuman(SourceText)
            .Execute Replace:=wdReplaceAll
        End With
    End With
End Sub

Private Function ConvertToOkuman(SourceText As String) As String
        
    'コンマ区切りから億万区切り
    Dim TargetChar As String
    
    '億万の付加
    SourceText = Format(SourceText, "####兆####億####万###0")
    
    Dim Length
    Length = Len(SourceText)
    
    Dim i As Long
    For i = 1 To Length
        If Left(SourceText, 1) Like "[兆,億,万]" Then
            SourceText = Right(SourceText, Length - i)
        Else
            Exit For
        End If
    Next
    
    '途中の0と次の区切りの除去
    If Right(SourceText, 4) = "0000" Then
        SourceText = Left(SourceText, Len(SourceText) - 4)
    End If
     
    If Right(SourceText, 5) = "0000万" Then
        SourceText = Left(SourceText, Len(SourceText) - 5)
    End If
    If Right(SourceText, 5) = "0000億" Then
        SourceText = Left(SourceText, Len(SourceText) - 5)
    End If
     
    '全角に変換
    ConvertToOkuman = StrConv(SourceText, vbWide)
     
End Function

Private Function ExpandRange(TargetRange As Range, CharSet As String) As Range
    '選択範囲の拡縮
    With TargetRange
        '選択範囲を広げる
        .MoveStartWhile CharSet, wdBackward
        .MoveEndWhile CharSet
        '選択範囲を狭める
        .MoveStartUntil CharSet
        .MoveEndUntil CharSet, wdBackward
    End With
    
    Set ExpandRange = TargetRange
    
End Function
