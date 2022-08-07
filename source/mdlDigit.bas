Attribute VB_Name = "mdlDigit"
'@Folder "main"
Option Explicit
Sub ArabicToChinese()
    '�A���r�A��������������
    
    '�I��͈͂��g�k����
    Dim TargetRange As Range
    Set TargetRange = ExpandRange(Selection.Range, "0123456789�O�P�Q�R�S�T�U�V�W�X�O,�C")
    
    '�I��͈͂̕����񂪋�Ȃ�I��
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
    '�������ϊ�
    '��S�̕t��
    SourceText = Format(SourceText, "#��#�S#�\0")
    Dim Length
    Length = Len(SourceText)
    
    Dim i As Long
    For i = 1 To Length
        If Left(SourceText, 1) Like "[��,�S,�\]" Then
            SourceText = Right(SourceText, Length - i)
        Else
            Exit For
        End If
    Next
    
    '�`���u1�v�̏���
    If Len(SourceText) > 1 And Left(SourceText, 1) = "1" Then
        SourceText = Right(SourceText, Len(SourceText) - 1)
    End If
    
    '�P�������ϊ�
    For i = 1 To 9
        SourceText = Replace(SourceText, CStr(i), Mid("���O�l�ܘZ������", i, 1))
    Next
    
     '�r����0�Ǝ��̋�؂�̏���
    If Right(SourceText, 1) = "0" Then
        SourceText = Left(SourceText, Len(SourceText) - 1)
    End If
    If Right(SourceText, 2) = "0�\" Then
        SourceText = Left(SourceText, Len(SourceText) - 2)
    End If
    If Right(SourceText, 2) = "0�S" Then
        SourceText = Left(SourceText, Len(SourceText) - 2)
    End If
     
    ConvertToChinese = SourceText
    
End Function

Sub ChineseToArabic()
    '���������A���r�A������
    
    '�I��͈͂��g�k����
    Dim TargetRange As Range
    Set TargetRange = ExpandRange(Selection.Range, "���O�l�ܘZ������\�S��")
    
    '�u��
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
    '�A���r�A�����ϊ�
    
    Dim i, j As Long
    For i = 1 To 9
        SourceText = Replace(SourceText, Mid("���O�l�ܘZ������", i, 1), CStr(i))
    Next
    
    SourceText = AddInitial(SourceText)

    Dim Splited() As String
    Dim Delimiter As String
    Dim Digits As String
    Dim DigitNumbers(3) As String
    Digits = "��S�\"
    
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
    If Left(SourceText, 1) = "��" Or Left(SourceText, 1) = "�S" Or Left(SourceText, 1) = "�\" Then
        AddInitial = "1" + SourceText
    Else
        AddInitial = SourceText
    End If
End Function

Sub OkumanToComma()
    '������؂���R���}��؂��
    
    '�I��͈͂��g�k����
    Dim TargetRange As Range
    Set TargetRange = ExpandRange(Selection.Range, "0123456789�O�P�Q�R�S�T�U�V�W�X�O������")
        
    '�u��
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
    '������؂肩��R���}��؂�
    
    Dim Splited() As String
    Dim Delimiter As String
    Dim Digits As String
    Dim DigitNumbers(3) As String
    Digits = "������"
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
    '�R���}��؂��������؂��
    
    '�I��͈͂��g�k����
    Dim TargetRange As Range
    Set TargetRange = ExpandRange(Selection.Range, "0123456789�O�P�Q�R�S�T�U�V�W�X�O,�C")
    
    '�I��͈͂̕����񂪋�Ȃ�I��
    If TargetRange.text = "" Then
        Exit Sub
    End If
    
    With TargetRange
        Dim SourceText As String
        SourceText = .text
        
        '�u��
        With .Find
            .text = SourceText
            .Replacement.text = ConvertToOkuman(SourceText)
            .Execute Replace:=wdReplaceAll
        End With
    End With
End Sub

Private Function ConvertToOkuman(SourceText As String) As String
        
    '�R���}��؂肩�牭����؂�
    Dim TargetChar As String
    
    '�����̕t��
    SourceText = Format(SourceText, "####��####��####��###0")
    
    Dim Length
    Length = Len(SourceText)
    
    Dim i As Long
    For i = 1 To Length
        If Left(SourceText, 1) Like "[��,��,��]" Then
            SourceText = Right(SourceText, Length - i)
        Else
            Exit For
        End If
    Next
    
    '�r����0�Ǝ��̋�؂�̏���
    If Right(SourceText, 4) = "0000" Then
        SourceText = Left(SourceText, Len(SourceText) - 4)
    End If
     
    If Right(SourceText, 5) = "0000��" Then
        SourceText = Left(SourceText, Len(SourceText) - 5)
    End If
    If Right(SourceText, 5) = "0000��" Then
        SourceText = Left(SourceText, Len(SourceText) - 5)
    End If
     
    '�S�p�ɕϊ�
    ConvertToOkuman = StrConv(SourceText, vbWide)
     
End Function

Private Function ExpandRange(TargetRange As Range, CharSet As String) As Range
    '�I��͈͂̊g�k
    With TargetRange
        '�I��͈͂��L����
        .MoveStartWhile CharSet, wdBackward
        .MoveEndWhile CharSet
        '�I��͈͂����߂�
        .MoveStartUntil CharSet
        .MoveEndUntil CharSet, wdBackward
    End With
    
    Set ExpandRange = TargetRange
    
End Function
