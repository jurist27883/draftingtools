Attribute VB_Name = "mdlDigit"
'@Folder "main"
Option Explicit

Sub ChangeDigit()
    Dim targetRange As Range
    Set targetRange = ExpandRange(Selection.Range, "0123456789０１２３４５６７８９０兆億万,，")
    
    If targetRange.text = "" Then
        Exit Sub
    End If
    
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord
    
    If InStr(targetRange.text, "兆") <> 0 Or InStr(targetRange.text, "億") _
        Or InStr(targetRange.text, "万") Then
        OkumanToComma targetRange
    Else
        CommaToOkuman targetRange
    End If
        
    ur.EndCustomRecord
End Sub

Sub ArabicToChinese()
    'アラビア数字を漢数字に
    
    '選択範囲を拡縮する
    Dim targetRange As Range
    Set targetRange = ExpandRange(Selection.Range, "0123456789０１２３４５６７８９０,，")
    
    '選択範囲の文字列が空なら終了
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
    '漢数字変換
    '千百の付加
    sourceText = Format(sourceText, "#千#百#十0")
    Dim length
    length = Len(sourceText)
    
    Dim i As Long
    For i = 1 To length
        If Left(sourceText, 1) Like "[千,百,十]" Then
            sourceText = Right(sourceText, length - i)
        Else
            Exit For
        End If
    Next
    
    '冒頭「1」の除去
    If Len(sourceText) > 1 And Left(sourceText, 1) = "1" Then
        sourceText = Right(sourceText, Len(sourceText) - 1)
    End If
    
    '単漢数字変換
    For i = 1 To 9
        sourceText = Replace(sourceText, CStr(i), Mid("一二三四五六七八九", i, 1))
    Next
    
     '途中の0と次の区切りの除去
    If Right(sourceText, 1) = "0" Then
        sourceText = Left(sourceText, Len(sourceText) - 1)
    End If
    If Right(sourceText, 2) = "0十" Then
        sourceText = Left(sourceText, Len(sourceText) - 2)
    End If
    If Right(sourceText, 2) = "0百" Then
        sourceText = Left(sourceText, Len(sourceText) - 2)
    End If
     
    ConvertToChinese = sourceText
    
End Function

Sub ChineseToArabic()
    '漢数字をアラビア数字に
    
    '選択範囲を拡縮する
    Dim targetRange As Range
    Set targetRange = ExpandRange(Selection.Range, "一二三四五六七八九十百千")
    
    '置換
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
    'アラビア数字変換
    
    Dim i, j As Long
    For i = 1 To 9
        sourceText = Replace(sourceText, Mid("一二三四五六七八九", i, 1), CStr(i))
    Next
    
    sourceText = AddInitial(sourceText)

    Dim splited() As String
    Dim delimiter As String
    Dim digits As String
    Dim digitNumbers(3) As String
    digits = "千百十"
    
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
    If Left(sourceText, 1) = "千" Or Left(sourceText, 1) = "百" Or Left(sourceText, 1) = "十" Then
        AddInitial = "1" + sourceText
    Else
        AddInitial = sourceText
    End If
End Function

Private Sub OkumanToComma(targetRange As Range)
    '億万区切りをコンマ区切りに
    
    '選択範囲を拡縮する
'    Dim targetRange As Range
'    Set targetRange = ExpandRange(Selection.Range, "0123456789０１２３４５６７８９０兆億万")
        
    '置換
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
    '億万区切りからコンマ区切り
    
    Dim splited() As String
    Dim delimiter As String
    Dim digits As String
    Dim digitNumbers(3) As String
    digits = "兆億万"
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
    'コンマ区切りを億万区切りに
    
    '選択範囲を拡縮する
'    Dim targetRange As Range
'    Set targetRange = ExpandRange(Selection.Range, "0123456789０１２３４５６７８９０,，")
    
    '選択範囲の文字列が空なら終了
    If targetRange.text = "" Then
        Exit Sub
    End If
    
    With targetRange
        Dim sourceText As String
        sourceText = .text
        
        '置換
        With .Find
            .text = sourceText
            .Replacement.text = ConvertToOkuman(sourceText)
            .Execute Replace:=wdReplaceAll
        End With
    End With
End Sub

Private Function ConvertToOkuman(sourceText As String) As String
        
    'コンマ区切りから億万区切り
    Dim TargetChar As String
    
    '億万の付加
    sourceText = Format(sourceText, "####兆####億####万###0")
    
    Dim length
    length = Len(sourceText)
    
    Dim i As Long
    For i = 1 To length
        If Left(sourceText, 1) Like "[兆,億,万]" Then
            sourceText = Right(sourceText, length - i)
        Else
            Exit For
        End If
    Next
    
    '途中の0と次の区切りの除去
    If Right(sourceText, 4) = "0000" Then
        sourceText = Left(sourceText, Len(sourceText) - 4)
    End If
     
    If Right(sourceText, 5) = "0000万" Then
        sourceText = Left(sourceText, Len(sourceText) - 5)
    End If
    If Right(sourceText, 5) = "0000億" Then
        sourceText = Left(sourceText, Len(sourceText) - 5)
    End If
     
    '全角に変換
    ConvertToOkuman = StrConv(sourceText, vbWide)
     
End Function

Private Function ExpandRange(targetRange As Range, charSet As String) As Range
    '選択範囲の拡縮
    With targetRange
        '選択範囲を広げる
        .MoveStartWhile charSet, wdBackward
        .MoveEndWhile charSet
        '選択範囲を狭める
        .MoveStartUntil charSet
        .MoveEndUntil charSet, wdBackward
    End With
    
    Set ExpandRange = targetRange
    
End Function
