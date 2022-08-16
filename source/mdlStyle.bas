Attribute VB_Name = "mdlStyle"
'@Folder "main.style"
Option Explicit

Public Const TITLE1 = "Title1"
Public Const TITLE2 = "Title2"
Public Const TITLE3 = "Title3"
Public Const TITLE4 = "Title4"
Public Const TITLE5 = "Title5"
Public Const BODY1 = "Body1"
Public Const BODY2 = "Body2"
Public Const BODY3 = "Body3"
Public Const BODY4 = "Body4"
Public Const BODY5 = "Body5"

Property Get IsGothic(styleName As String) As Boolean
    '�S�V�b�N����
    Select Case ActiveDocument.Styles(styleName).Font.NameFarEast
    Case MSGOTHIC, YUGOTHIC, BIZUDGOTHIC
        IsGothic = True
    End Select
End Property

Property Get IsBold(styleName As String) As Boolean
    '��������
    If ActiveDocument.Styles(styleName).Font.Bold Then
        IsBold = True
    End If
End Property

Sub ToggleFontFamily(styleName As String)
    '�����E�S�V�b�N�ؑւ�
    With ActiveDocument.Styles(styleName).Font
        .NameFarEast = CounterFontName(ActiveDocument.Styles(styleName).Font.NameFarEast)
        .Name = .NameFarEast
    End With
End Sub

Sub SetTitle1()
    SetStyleToParagragh TITLE1
End Sub
Sub SetTitle2()
    SetStyleToParagragh TITLE2
End Sub
Sub SetTitle3()
    SetStyleToParagragh TITLE3
End Sub
Sub SetTitle4()
    SetStyleToParagragh TITLE4
End Sub
Sub SetTitle5()
    SetStyleToParagragh TITLE5
End Sub
Sub SetBody1()
    SetStyleToParagragh BODY1
End Sub
Sub SetBody2()
    SetStyleToParagragh BODY2
End Sub
Sub SetBody3()
    SetStyleToParagragh BODY3
End Sub
Sub SetBody4()
    SetStyleToParagragh BODY4
End Sub
Sub SetBody5()
    SetStyleToParagragh BODY5
End Sub

Private Sub SetStyleToParagragh(styleName As String)
    '�X�^�C���K�p
    If Documents.Count = 0 Then
        Exit Sub
    End If
    
    ResisterStyle styleName
    
    Dim p As Paragraph
    For Each p In Selection.Paragraphs
        p.Style = ActiveDocument.Styles(styleName)
    Next
End Sub

Sub CopyFontFromNormalTemplete(styleName As String)
    '�X�^�C���ɕW���e���v���[�g�̃t�H���g��ݒ�
    Dim gt As Boolean
    If IsGothic(styleName) Then
        gt = True
    End If
    
    With ActiveDocument.Styles(styleName).Font
        .NameFarEast = ActiveDocument.Styles(wdStyleNormal).Font.NameFarEast
        .Name = ActiveDocument.Styles(wdStyleNormal).Font.Name
        If gt Then
            ToggleFontFamily styleName
        End If
        .Size = ActiveDocument.Styles(wdStyleNormal).Font.Size
    End With
    
End Sub

Sub SetSavedStylesFromXml(styleName As String)
    'xml�ۑ��X�^�C����K�p
    Dim willGothic As Variant
    Dim willBold As Variant
    
    For Each willGothic In mdlXml.SelectedTexts(TAG_STYLE, TAG_STYLE_NAME, styleName, TAG_GOTHIC)
        If CBool(willGothic) <> IsGothic(styleName) Then
            ToggleFontFamily styleName
        End If
    Next
    
    For Each willBold In mdlXml.SelectedTexts(TAG_STYLE, TAG_STYLE_NAME, styleName, TAG_BOLD)
        ActiveDocument.Styles(styleName).Font.Bold = CBool(willGothic)
    Next
End Sub

Private Sub CopyStyleFromDraftingTemplete(styleName As String)
    '�X�^�C�����N�ăe���v���[�g����R�s�[
    Application.OrganizerCopy ThisDocument.FullName, ActiveDocument.FullName, _
                              styleName, wdOrganizerObjectStyles
End Sub

Sub ResisterStyle(styleName As String)
    '�X�^�C���̗L���`�F�b�N�Ƒ��݂��Ȃ��ꍇ�̃R�s�[
    Dim dummy As String
    On Error GoTo Catch
    dummy = ActiveDocument.Styles(styleName).Font.Name
    Exit Sub
Catch:
    If Err.Number = 5941 Then
        CopyStyleFromDraftingTemplete styleName
        SetSavedStylesFromXml styleName
        CopyFontFromNormalTemplete styleName
        Resume
    Else
        MsgBox "�G���[" + Err.Number + "���������܂����B"
    End If
End Sub

Sub IncreaseOutlineLevel()
    ' �A�E�g���C�����x������
    'todo ����
    Dim p As Paragraph
    For Each p In Selection.Paragraphs
        With p
            Select Case .OutlineLevel
            Case wdOutlineLevelBodyText
                .OutlineLevel = 1
            Case 1 To 8
                .OutlineLevel = .OutlineLevel + 1
            End Select
        End With
    Next
End Sub
Sub DecreaseOutlineLevel()
    ' �A�E�g���C�����x������
    'todo ����
    Dim p As Paragraph
    For Each p In Selection.Paragraphs
        With p
            Select Case .OutlineLevel
            Case 1
                .OutlineLevel = wdOutlineLevelBodyText
            Case 2 To 9
                .OutlineLevel = .OutlineLevel - 1
            End Select
        End With
    Next
End Sub

Sub ClearOutlineLevel()
    ' �A�E�g���C�����x���폜
    'todo ����
    With Selection.ParagraphFormat
        .OutlineLevel = wdOutlineLevelBodyText
    End With
End Sub

Sub ChangeStyle()
    '�X�^�C������������
    'todo ����
    Dim p As Paragraph
    Dim styleString As String
    For Each p In Selection.Paragraphs
        styleString = p.Style
        Select Case styleString
        Case "Title5"
            p.Style = "Body5"
        Case "Title4"
            p.Style = "Body4"
        Case "Title3"
            p.Style = "Body3"
        Case "Title2"
            p.Style = "Body2"
        Case "Title1"
            p.Style = "Body1"
        Case "Body5"
            p.Style = "Title5"
        Case "Body4"
            p.Style = "Title4"
        Case "Body3"
            p.Style = "Title3"
        Case "Body2"
            p.Style = "Title2"
        Case "Body1"
            p.Style = "Title1"
        End Select
    Next
End Sub

Sub IncleaseBodyStyle()
    '�{���X�^�C���̃��x�����グ��
    'todo ����
    Dim p As Paragraph
    Dim styleString As String
    For Each p In Selection.Paragraphs
        styleString = p.Style
        Select Case styleString
        Case "Body5"
            p.Style = "Body4"
        Case "Body4"
            p.Style = "Body3"
        Case "Body3"
            p.Style = "Body2"
        Case "Body2"
            p.Style = "Body1"
        Case Else
            p.Style = "Body1"
        End Select
    Next
End Sub

Sub DecleaseBodyStyle()
    '�{���X�^�C���̃��x����������
    'todo ����
    Dim p As Paragraph
    Dim styleString As String
    
    For Each p In Selection.Paragraphs
    
        styleString = p.Style
        
        Select Case styleString
        Case "Body1"
            p.Style = "Body2"
        Case "Body2"
            p.Style = "Body3"
        Case "Body3"
            p.Style = "Body4"
        Case "Body4"
            p.Style = "Body5"
        Case Else
            p.Style = "Body1"
        End Select
    Next
End Sub

Sub IncleaseTitleStyle()
    '�^�C�g���X�^�C���̃��x�����グ��
    'todo ����
    Dim p As Paragraph
    Dim styleString As String
    
    For Each p In Selection.Paragraphs
    
        styleString = p.Style
        
        Select Case styleString
        Case "Title5"
            p.Style = "Title4"
        Case "Title4"
            p.Style = "Title3"
        Case "Title3"
            p.Style = "Title2"
        Case "Title2"
            p.Style = "Title1"
        Case Else
            p.Style = "Title1"
        End Select
    Next
End Sub

Sub DecleaseTitleStyle()
    '�^�C�g���X�^�C���̃��x����������
    'todo ����
        
    Dim p As Paragraph
    Dim styleString As String
    
    For Each p In Selection.Paragraphs
        styleString = p.Style
        
        Select Case styleString
        Case "Title1"
            p.Style = "Title2"
        Case "Title2"
            p.Style = "Title3"
        Case "Title3"
            p.Style = "Title4"
        Case "Title4"
            p.Style = "Title5"
        Case Else
            p.Style = "Title1"
        End Select
    Next
End Sub

Sub ClearStyle()
    '�X�^�C���̃N���A
    Dim p As Paragraph
    For Each p In Selection.Paragraphs
        p.Style = ActiveDocument.Styles(wdStyleNormal)
    Next
End Sub


