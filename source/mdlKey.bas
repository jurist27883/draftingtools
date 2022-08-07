Attribute VB_Name = "mdlKey"
'@Folder "keycode"
Option Explicit

Property Get BindingCommand(code As Long) As String
    '�L�[�R�[�h����o�^���ꂽ�R�}���h���擾
    Dim kb As KeyBinding
    Set kb = GetBindingByCode(code)
    If Not (kb Is Nothing) Then
        BindingCommand = kb.command
    End If
End Property

Property Get BindingCode(targetCommand As String) As Long
    '�R�}���h����o�^���ꂽ�L�[�R�[�h���擾
    Dim kb As KeyBinding
    Set kb = GetBindingByCommand(targetCommand)
    If Not (kb Is Nothing) Then
        BindingCode = kb.KeyCode
    End If
End Property

Property Let BindingCode(targetCommand As String, assigningKeyCode As Long)
    '�V���[�g�J�b�g�L�[�̓o�^
    #If RELEASE Then
        CustomizationContext = NormalTemplate
    #Else
        CustomizationContext = ThisDocument
    #End If
    
    KeyBindings.Add wdKeyCategoryMacro, targetCommand, assigningKeyCode
End Property

Private Function GetBindingByCode(code As Long) As KeyBinding
    '�L�[�R�[�h����Keybinding���擾
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
    '�R�}���h����Keybinding���擾
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
    '�V���[�g�J�b�g�L�[����
    If frmConfig.lstKeys.ListIndex = -1 Then
        Exit Sub
    End If
    If mdlListBoxKeys.SelectedItem(CL.BINDING_KEY_CODE) = 0 Then
        Exit Sub
    End If
    
    '�L�[�o�C���f�B���O����
    ResetKeyBindingByCommand command
    
    '���̃V���[�g�J�b�g�L�[�̕�������/config�t�@�C������̍폜
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

