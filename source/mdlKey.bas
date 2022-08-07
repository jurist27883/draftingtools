Attribute VB_Name = "mdlKey"
'@Folder "keycode"
Option Explicit

Property Get BindingCommand(code As Long) As String
    'キーコードから登録されたコマンドを取得
    Dim kb As KeyBinding
    Set kb = GetBindingByCode(code)
    If Not (kb Is Nothing) Then
        BindingCommand = kb.command
    End If
End Property

Property Get BindingCode(targetCommand As String) As Long
    'コマンドから登録されたキーコードを取得
    Dim kb As KeyBinding
    Set kb = GetBindingByCommand(targetCommand)
    If Not (kb Is Nothing) Then
        BindingCode = kb.KeyCode
    End If
End Property

Property Let BindingCode(targetCommand As String, assigningKeyCode As Long)
    'ショートカットキーの登録
    #If RELEASE Then
        CustomizationContext = NormalTemplate
    #Else
        CustomizationContext = ThisDocument
    #End If
    
    KeyBindings.Add wdKeyCategoryMacro, targetCommand, assigningKeyCode
End Property

Private Function GetBindingByCode(code As Long) As KeyBinding
    'キーコードからKeybindingを取得
    #If RELEASE Then
        CustomizationContext = NormalTemplate
    #Else
        CustomizationContext = ThisDocument
    #End If
    
    Dim kb As KeyBinding
    For Each kb In KeyBindings
        If kb.KeyCode = code Then
            Set GetBindingByCode = kb
        End If
    Next
End Function

Private Function GetBindingByCommand(targetCommand As String) As KeyBinding
    'コマンドからKeybindingを取得
    #If RELEASE Then
        CustomizationContext = NormalTemplate
    #Else
        CustomizationContext = ThisDocument
    #End If
    
    Dim kb As KeyBinding
    For Each kb In KeyBindings
        If kb.command = targetCommand Then
            Set GetBindingByCommand = kb
        End If
    Next
End Function

Sub ResetKey(command As String, KeyCode As Long)
    'ショートカットキー解除
    If frmConfig.lstKeys.ListIndex = -1 Then
        Exit Sub
    End If
    If mdlListBoxKeys.SelectedItem(CL.BINDING_KEY_CODE) = 0 Then
        Exit Sub
    End If
    
    'キーバインディング解除
    ResetKeyBindingByCommand command
    
    '元のショートカットキーの復活処理/configファイルからの削除
    Dim cmd As Variant
    ReresisterKeyBinding KeyCode
End Sub

Sub ReresisterKeyBinding(KeyCode As Long)
    Dim cmd As Variant
    For Each cmd In mdlXml.SelectedTexts(TAG_KEYBINDING, TAG_KEYCODE, CStr(KeyCode), TAG_FORMER_COMMAND)
        If CStr(cmd) <> "" Then
            mdlKey.BindingCode(CStr(cmd)) = KeyCode
        End If
        mdlXml.RemoveNode TAG_KEYBINDING, TAG_KEYCODE, CStr(KeyCode)
    Next
End Sub

Sub ResetKeyBindingByCommand(command As String)
    #If RELEASE Then
        CustomizationContext = NormalTemplate
    #Else
        CustomizationContext = ThisDocument
    #End If
    Dim kb As KeyBinding
    Set kb = GetBindingByCommand(command)
    If Not (kb Is Nothing) Then
        kb.Clear
    End If
End Sub

