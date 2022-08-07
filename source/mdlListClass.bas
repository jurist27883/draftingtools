Attribute VB_Name = "mdlListBoxKeys"
'@Folder("config")
Option Explicit

Enum CL
    COMMAND_NAME = 0
    COMMAND_STRING = 1
    BINDING_KEY_CODE = 2
    BINDING_KEY_STRING = 3
    RECOMMEND_KEY_CODE = 4
End Enum

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

Private Sub AddRow(lst() As Variant, ByRef rw As Long, commandName As String, commandString As String, _
                   recommendKeyCode)
    'リストボックス行追加
    lst(rw, COMMAND_NAME) = commandName
    lst(rw, COMMAND_STRING) = commandString
    lst(rw, CL.BINDING_KEY_CODE) = mdlKey.BindingCode(commandString)
    If lst(rw, CL.BINDING_KEY_CODE) <> 0 Then
        lst(rw, CL.BINDING_KEY_STRING) = KeyString(mdlKey.BindingCode(commandString))
    End If
    lst(rw, CL.RECOMMEND_KEY_CODE) = recommendKeyCode
    
    rw = rw + 1
End Sub

Sub Initialize()
    'リストボックス初期化
    Const columnsCount = 5
    Const rowsCount = 18
    Dim lst(rowsCount - 1, columnsCount - 1) As Variant
    
    Dim rw As Long
    AddRow lst, rw, COMMAND_NAME_RIGHT, COMMAND_RIGHT, RECOMMEND.KEY_RIGHT
    AddRow lst, rw, COMMAND_NAME_LEFT, COMMAND_LEFT, RECOMMEND.KEY_LEFT
    AddRow lst, rw, COMMAND_NAME_PRIMARY_RIGHT, COMMAND_PRIMARY_RIGHT, RECOMMEND.KEY_PRIMARY_RIGHT
    AddRow lst, rw, COMMAND_NAME_PRIMARY_LEFT, COMMAND_PRIMARY_LEFT, RECOMMEND.KEY_PRIMARY_LEFT
    AddRow lst, rw, COMMAND_NAME_SECONDARY_RIGHT, COMMAND_SECONDARY_RIGHT, RECOMMEND.KEY_SECONDARY_RIGHT
    AddRow lst, rw, COMMAND_NAME_SECONDARY_LEFT, COMMAND_SECONDARY_LEFT, RECOMMEND.KEY_SECONDARY_LEFT
    AddRow lst, rw, COMMAND_NAME_ROUND, COMMAND_ROUND, RECOMMEND.KEY_ROUND
    AddRow lst, rw, COMMAND_NAME_TITLE1, COMMAND_TITLE1, RECOMMEND.KEY_TITLE1
    AddRow lst, rw, COMMAND_NAME_TITLE2, COMMAND_TITLE2, RECOMMEND.KEY_TITLE2
    AddRow lst, rw, COMMAND_NAME_TITLE3, COMMAND_TITLE3, RECOMMEND.KEY_TITLE3
    AddRow lst, rw, COMMAND_NAME_TITLE4, COMMAND_TITLE4, RECOMMEND.KEY_TITLE4
    AddRow lst, rw, COMMAND_NAME_TITLE5, COMMAND_TITLE5, RECOMMEND.KEY_TITLE5
    AddRow lst, rw, COMMAND_NAME_BODY1, COMMAND_BODY1, RECOMMEND.KEY_BODY1
    AddRow lst, rw, COMMAND_NAME_BODY2, COMMAND_BODY2, RECOMMEND.KEY_BODY2
    AddRow lst, rw, COMMAND_NAME_BODY3, COMMAND_BODY3, RECOMMEND.KEY_BODY3
    AddRow lst, rw, COMMAND_NAME_BODY4, COMMAND_BODY4, RECOMMEND.KEY_BODY4
    AddRow lst, rw, COMMAND_NAME_BODY5, COMMAND_BODY5, RECOMMEND.KEY_BODY5
    AddRow lst, rw, COMMAND_NAME_CLEAR_STYLE, COMMAND_CLEAR_STYLE, RECOMMEND.KEY_CLEAR_STYLE
    
    With frmConfig.lstKeys
        .ColumnCount = columnsCount
        .ColumnWidths = ";0;0;;0"
        .List = lst
    End With
End Sub


