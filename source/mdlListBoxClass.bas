Attribute VB_Name = "mdlListBoxClass"
'@Folder "config.form"
Option Explicit

Enum CLASS_INDEX
    INSDENT_NUMBER = 0
    STYLE_NUMBER = 1
End Enum

Const CLASS_STRING_INDENT = "�C���f���g"
Const CLASS_STRING_STYLE = "�i���X�^�C��"

Property Get SelectedItem(column As Long) As Variant
    If frmConfig.lstKeys.ListIndex = -1 Then
        SelectedItem = Nothing
    Else
        SelectedItem = frmConfig.lstKeys.List(frmConfig.lstKeys.ListIndex, column)
    End If
End Property

Property Let SelectedItem(column As Long, value As Variant)
    If frmConfig.lstKeys.ListIndex = -1 Then
        Exit Property
    Else
        frmConfig.lstKeys.List(frmConfig.lstKeys.ListIndex, column) = value
    End If
End Property

Private Sub AddRow(lst() As Variant, ByRef rw As Long, classString As String)
    '���X�g�{�b�N�X�s�ǉ�
    lst(rw, 0) = classString
    
    rw = rw + 1
End Sub

Sub Initialize()
    '���X�g�{�b�N�X������
    Const columnsCount = 1
    Const rowsCount = 2
    Dim lst(rowsCount - 1, columnsCount - 1) As Variant
    
    Dim rw As Long
    AddRow lst, rw, CLASS_STRING_INDENT
    AddRow lst, rw, CLASS_STRING_STYLE
    
    With frmConfig.lstClass
        .ColumnCount = columnsCount
        .List = lst
    End With
End Sub

