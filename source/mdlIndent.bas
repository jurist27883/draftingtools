Attribute VB_Name = "mdlIndent"
'@Folder "main"
Option Explicit
Sub IndentPrimaryRight()
    '字下げ右インデント
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
    '字下げ右インデントメソッド
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
    '字下げ左インデント
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
    '字下げ左インデントメソッド
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
    'ぶら下げ右インデント
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
    'ぶら下げ右インデントメソッド
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
    'ぶら下げ左インデント
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
    'ぶら下げ左インデントメソッド
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
    '選択範囲の全段落の全体を右にインデント
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
    '全体右インデントメソッド
    AdjustNotCharSystem targetParagragh
    
    With targetParagragh.Format
        '右端超過防止
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
    '選択範囲の全段落の全体を左にインデント
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
    '全体左インデントメソッド
    AdjustNotCharSystem targetParagragh
    
    With targetParagragh.Format
        If .CharacterUnitLeftIndent - 1 > 0 Then
            .CharacterUnitLeftIndent = .CharacterUnitLeftIndent - 1
        Else
            'char系・非char系調整
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
    '非char系にのみ数値が設定されている場合への対応
    With targetParagraph.Format
        If (.CharacterUnitLeftIndent = 0 And .LeftIndent <> 0) _
        Or (.CharacterUnitFirstLineIndent = 0 And .FirstLineIndent <> 0) Then
            IndentRound
        End If
    End With
End Sub

Sub IndentFirstLineRight()
    '選択範囲の全段落の字下げ・ぶら下げインデントを右に増やす
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
    '字下げ・ぶら下げ右インデントメソッド
    AdjustNotCharSystem targetParagragh
    
    With targetParagragh.Format
        Select Case .CharacterUnitFirstLineIndent + 1
        Case Is > 0
            '右端超過防止
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
    '選択範囲の全段落の字下げ・ぶら下げインデントを左に減らす
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
    '字下げ・ぶら下げ左インデントメソッド
    AdjustNotCharSystem targetParagragh
    
    With targetParagragh.Format
        Select Case .CharacterUnitFirstLineIndent - 1
        Case Is > 0
            .CharacterUnitFirstLineIndent = .CharacterUnitFirstLineIndent - 1
        Case 0
            .CharacterUnitFirstLineIndent = 0
            .FirstLineIndent = 0
        Case Is < 0
            '右端超過防止
            If targetParagragh.Parent.PageSetup.CharsLine - 1 < _
            .CharacterUnitLeftIndent - (.CharacterUnitFirstLineIndent - 1) Then
                Exit Sub
            End If
            
            .CharacterUnitFirstLineIndent = .CharacterUnitFirstLineIndent - 1
        End Select
    End With
End Sub

Sub IndentRound()
    '選択範囲の全段落のインデントを整数値化
    If Documents.Count = 0 Then
        Exit Sub
    End If
    
    Dim pr As Paragraph
    For Each pr In Selection.Paragraphs
        ExecRound pr
    Next
End Sub

Private Sub ExecRound(targetParagragh As Paragraph)
    'インデント整数値化メソッド
     With targetParagragh.Format
        '字送り取得
        Dim charWidth As Long
        With targetParagragh.Range.PageSetup
            charWidth = (targetParagragh.Parent.PageSetup.PageWidth - (.LeftMargin + .RightMargin)) / .CharsLine
        End With
        
        '整数に変換
        Dim firstLinecharIndent As Long
        firstLinecharIndent = Format(.FirstLineIndent / charWidth, 0)
        
        Dim charIndent As Long
        If .FirstLineIndent >= 0 Then
            charIndent = Format(.LeftIndent / charWidth, 0)
        Else
            charIndent = Format((.FirstLineIndent + .LeftIndent) / charWidth, 0)
        End If
                
        'Char系に0を設定する場合、非Char系に0以外が設定されていると実際のインデントに反映されないことや
        '後にChar系を0にしたときに非Char系が復活することへの対応
        .FirstLineIndent = 0
        .LeftIndent = 0
        
        '整数文字インデント設定
        .CharacterUnitFirstLineIndent = firstLinecharIndent
        .CharacterUnitLeftIndent = charIndent
    End With
End Sub


