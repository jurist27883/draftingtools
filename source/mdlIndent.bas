Attribute VB_Name = "mdlIndent"
'@Folder "main"
Option Explicit
Sub IndentPrimaryRight()
    '�������E�C���f���g
    If Documents.Count = 0 Then
        Exit Sub
    End If
    
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord
    
    Dim pr As Paragraph
    For Each pr In Selection.Paragraphs
        ExecPrimaryRight pr
    Next

    ur.EndCustomRecord
End Sub

Private Sub ExecPrimaryRight(targetParagragh As Paragraph)
    '�������E�C���f���g���\�b�h
    AdjustNotCharSystem targetParagragh
    
    With targetParagragh.Format
        If .CharacterUnitFirstLineIndent >= 0 Then
            ExecFirstRight targetParagragh
        Else
            Application.ScreenUpdating = False
            ExecFirstRight targetParagragh
            ExecRight targetParagragh
            Application.ScreenUpdating = True
        End If
    End With
End Sub

Sub IndentPrimaryLeft()
    '���������C���f���g
    If Documents.Count = 0 Then
        Exit Sub
    End If
    
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord
    
    Dim pr As Paragraph
    For Each pr In Selection.Paragraphs
        ExecPrimaryLeft pr
    Next

    ur.EndCustomRecord
End Sub

Private Sub ExecPrimaryLeft(targetParagragh As Paragraph)
    '���������C���f���g���\�b�h
    AdjustNotCharSystem targetParagragh
    
    With targetParagragh.Format
        If .CharacterUnitFirstLineIndent > 0 Then
            ExecFirstLeft targetParagragh
        Else
            If .CharacterUnitLeftIndent > 0 Then
                Application.ScreenUpdating = False
                ExecFirstLeft targetParagragh
                ExecLeft targetParagragh
                Application.ScreenUpdating = True
            End If
        End If
    End With
End Sub

Sub IndentSecondaryRight()
    '�Ԃ牺���E�C���f���g
    If Documents.Count = 0 Then
        Exit Sub
    End If
    
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord
    
    Dim pr As Paragraph
    For Each pr In Selection.Paragraphs
        ExecSecondaryRight pr
    Next

    ur.EndCustomRecord
End Sub

Private Sub ExecSecondaryRight(targetParagragh As Paragraph)
    '�Ԃ牺���E�C���f���g���\�b�h
    AdjustNotCharSystem targetParagragh
    
    With targetParagragh.Format
        If .CharacterUnitFirstLineIndent <= 0 Then
            ExecFirstLeft targetParagragh
        Else
            Application.ScreenUpdating = False
            ExecFirstLeft targetParagragh
            ExecRight targetParagragh
            Application.ScreenUpdating = True
        End If
    End With
End Sub

Sub IndentSecondaryLeft()
    '�Ԃ牺�����C���f���g
    If Documents.Count = 0 Then
        Exit Sub
    End If
    
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord
    
    Dim pr As Paragraph
    For Each pr In Selection.Paragraphs
        ExecSecondaryLeft pr
    Next

    ur.EndCustomRecord
End Sub

Private Sub ExecSecondaryLeft(targetParagragh As Paragraph)
    '�Ԃ牺�����C���f���g���\�b�h
    AdjustNotCharSystem targetParagragh
    
    With targetParagragh.Format
        If .CharacterUnitFirstLineIndent < 0 Then
            ExecFirstRight targetParagragh
        Else
            If .CharacterUnitLeftIndent > 0 Then
                Application.ScreenUpdating = False
                ExecFirstRight targetParagragh
                ExecLeft targetParagragh
                Application.ScreenUpdating = True
            End If
        End If
    End With
End Sub

Sub IndentRight()
    '�I��͈͂̑S�i���̑S�̂��E�ɃC���f���g
    If Documents.Count = 0 Then
        Exit Sub
    End If
    
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord
    
    Dim pr As Paragraph
    For Each pr In Selection.Paragraphs
        ExecRight pr
    Next
    
    ur.EndCustomRecord
End Sub

Private Sub ExecRight(targetParagragh As Paragraph)
    '�S�̉E�C���f���g���\�b�h
    AdjustNotCharSystem targetParagragh
    
    With targetParagragh.Format
        '�E�[���ߖh�~
        If .CharacterUnitFirstLineIndent >= 0 Then
            If targetParagragh.Parent.PageSetup.CharsLine - 1 < _
            .CharacterUnitLeftIndent + .CharacterUnitFirstLineIndent + 1 Then
                Exit Sub
            End If
        Else
            If targetParagragh.Parent.PageSetup.CharsLine - 1 < _
            .CharacterUnitLeftIndent - (.CharacterUnitFirstLineIndent - 1) Then
                Exit Sub
            End If
        End If
        
        .CharacterUnitLeftIndent = .CharacterUnitLeftIndent + 1
    End With
End Sub

Sub IndentLeft()
    '�I��͈͂̑S�i���̑S�̂����ɃC���f���g
    If Documents.Count = 0 Then
        Exit Sub
    End If

    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord
    
    Dim pr As Paragraph
    For Each pr In Selection.Paragraphs
        ExecLeft pr
    Next

    ur.EndCustomRecord
End Sub

Private Sub ExecLeft(targetParagragh As Paragraph)
    '�S�̍��C���f���g���\�b�h
    AdjustNotCharSystem targetParagragh
    
    With targetParagragh.Format
        If .CharacterUnitLeftIndent - 1 > 0 Then
            .CharacterUnitLeftIndent = .CharacterUnitLeftIndent - 1
        Else
            'char�n�E��char�n����
            If .CharacterUnitFirstLineIndent >= 0 Then
                .CharacterUnitLeftIndent = 0
                .LeftIndent = 0
            Else
                .CharacterUnitLeftIndent = 0
                .FirstLineIndent = 0
            End If
        End If
    End With
End Sub

Private Sub AdjustNotCharSystem(targetParagraph As Paragraph)
    '��char�n�ɂ̂ݐ��l���ݒ肳��Ă���ꍇ�ւ̑Ή�
    With targetParagraph.Format
        If (.CharacterUnitLeftIndent = 0 And .LeftIndent <> 0) _
        Or (.CharacterUnitFirstLineIndent = 0 And .FirstLineIndent <> 0) Then
            IndentRound
        End If
    End With
End Sub

Sub IndentFirstLineRight()
    '�I��͈͂̑S�i���̎������E�Ԃ牺���C���f���g���E�ɑ��₷
    If Documents.Count = 0 Then
        Exit Sub
    End If
    
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord
    
    Dim pr As Paragraph
    For Each pr In Selection.Paragraphs
        ExecFirstRight pr
    Next

    ur.EndCustomRecord
End Sub

Private Sub ExecFirstRight(targetParagragh As Paragraph)
    '�������E�Ԃ牺���E�C���f���g���\�b�h
    AdjustNotCharSystem targetParagragh
    
    With targetParagragh.Format
        Select Case .CharacterUnitFirstLineIndent + 1
        Case Is > 0
            '�E�[���ߖh�~
            If targetParagragh.Parent.PageSetup.CharsLine - 1 < _
            .CharacterUnitLeftIndent + .CharacterUnitFirstLineIndent + 1 Then
                Exit Sub
            End If
            
            .CharacterUnitFirstLineIndent = .CharacterUnitFirstLineIndent + 1
        Case 0
            .CharacterUnitFirstLineIndent = 0
            .FirstLineIndent = 0
            If .CharacterUnitLeftIndent = 0 Then
                .LeftIndent = 0
            End If
        Case Is < 0
            .CharacterUnitFirstLineIndent = .CharacterUnitFirstLineIndent + 1
        End Select
    End With
End Sub

Sub IndentFirstLineLeft()
    '�I��͈͂̑S�i���̎������E�Ԃ牺���C���f���g�����Ɍ��炷
    If Documents.Count = 0 Then
        Exit Sub
    End If
    
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord
    
    Dim pr As Paragraph
    For Each pr In Selection.Paragraphs
        ExecFirstLeft pr
    Next
    
    ur.EndCustomRecord
End Sub

Private Sub ExecFirstLeft(targetParagragh As Paragraph)
    '�������E�Ԃ牺�����C���f���g���\�b�h
    AdjustNotCharSystem targetParagragh
    
    With targetParagragh.Format
        Select Case .CharacterUnitFirstLineIndent - 1
        Case Is > 0
            .CharacterUnitFirstLineIndent = .CharacterUnitFirstLineIndent - 1
        Case 0
            .CharacterUnitFirstLineIndent = 0
            .FirstLineIndent = 0
        Case Is < 0
            '�E�[���ߖh�~
            If targetParagragh.Parent.PageSetup.CharsLine - 1 < _
            .CharacterUnitLeftIndent - (.CharacterUnitFirstLineIndent - 1) Then
                Exit Sub
            End If
            
            .CharacterUnitFirstLineIndent = .CharacterUnitFirstLineIndent - 1
        End Select
    End With
End Sub

Sub IndentRound()
    '�I��͈͂̑S�i���̃C���f���g�𐮐��l��
    If Documents.Count = 0 Then
        Exit Sub
    End If
    
    Dim pr As Paragraph
    For Each pr In Selection.Paragraphs
        ExecRound pr
    Next
End Sub

Private Sub ExecRound(targetParagragh As Paragraph)
    '�C���f���g�����l�����\�b�h
     With targetParagragh.Format
        '������擾
        Dim charWidth As Long
        With targetParagragh.Range.PageSetup
            charWidth = (targetParagragh.Parent.PageSetup.PageWidth - (.LeftMargin + .RightMargin)) / .CharsLine
        End With
        
        '�����ɕϊ�
        Dim firstLinecharIndent As Long
        firstLinecharIndent = Format(.FirstLineIndent / charWidth, 0)
        
        Dim charIndent As Long
        If .FirstLineIndent >= 0 Then
            charIndent = Format(.LeftIndent / charWidth, 0)
        Else
            charIndent = Format((.FirstLineIndent + .LeftIndent) / charWidth, 0)
        End If
                
        'Char�n��0��ݒ肷��ꍇ�A��Char�n��0�ȊO���ݒ肳��Ă���Ǝ��ۂ̃C���f���g�ɔ��f����Ȃ����Ƃ�
        '���Char�n��0�ɂ����Ƃ��ɔ�Char�n���������邱�Ƃւ̑Ή�
        .FirstLineIndent = 0
        .LeftIndent = 0
        
        '���������C���f���g�ݒ�
        .CharacterUnitFirstLineIndent = firstLinecharIndent
        .CharacterUnitLeftIndent = charIndent
    End With
End Sub


