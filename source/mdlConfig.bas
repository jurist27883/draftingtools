Attribute VB_Name = "mdlConfig"
'@Folder("config")
Option Explicit

Public assigningKeyCode As Long

Public Enum MASK
    SHIFTS = 1
    CTRLS = 2
    ALTS = 4
End Enum

Public Const CONFIG_FILE_NAME = "draftingtools.xml"
Public Const TAG_KEYBINDING = "keybinding"
Public Const TAG_KEYCODE = "keycode"
Public Const TAG_FORMER_COMMAND = "formercommand"
Public Const TAG_STYLE = "style"
Public Const TAG_STYLE_NAME = "stylename"
Public Const TAG_GOTHIC = "gothic"
Public Const TAG_BOLD = "bold"

Public Const CSIDL_FONTS = 20

Public Const COMMAND_NAME_RIGHT = "�S�̃C���f���g �E"
Public Const COMMAND_NAME_LEFT = "�S�̃C���f���g ��"
Public Const COMMAND_NAME_PRIMARY_RIGHT = "1�s�ڃC���f���g �E"
Public Const COMMAND_NAME_PRIMARY_LEFT = "1�s�ڃC���f���g �E"
Public Const COMMAND_NAME_SECONDARY_RIGHT = "2�s�ڈȉ��C���f���g �E"
Public Const COMMAND_NAME_SECONDARY_LEFT = "2�s�ڈȉ��C���f���g �E"
Public Const COMMAND_NAME_ROUND = "�C���f���g�l�ۂ�"
Public Const COMMAND_NAME_TITLE1 = "���o�� [��1]"
Public Const COMMAND_NAME_TITLE2 = "���o�� [1]"
Public Const COMMAND_NAME_TITLE3 = "���o�� [(1)]"
Public Const COMMAND_NAME_TITLE4 = "���o�� [�A]"
Public Const COMMAND_NAME_TITLE5 = "���o�� [(�A)]"
Public Const COMMAND_NAME_BODY1 = "�{��  [��1]"
Public Const COMMAND_NAME_BODY2 = "�{��  [1]"
Public Const COMMAND_NAME_BODY3 = "�{��  [(1)]"
Public Const COMMAND_NAME_BODY4 = "�{��  [�A]"
Public Const COMMAND_NAME_BODY5 = "�{��  [(�A)]"
Public Const COMMAND_NAME_CLEAR_STYLE = "�N���A"

Public Const COMMAND_RIGHT = "Drafting.mdlIndent.IndentRight"
Public Const COMMAND_LEFT = "Drafting.mdlIndent.IndentLeft"
Public Const COMMAND_PRIMARY_RIGHT = "Drafting.mdlIndent.IndentPrimaryRight"
Public Const COMMAND_PRIMARY_LEFT = "Drafting.mdlIndent.IndentPrimaryLeft"
Public Const COMMAND_SECONDARY_RIGHT = "Drafting.mdlIndent.IndentSecondaryRight"
Public Const COMMAND_SECONDARY_LEFT = "Drafting.mdlIndent.IndentSecondaryLeft"
Public Const COMMAND_ROUND = "Drafting.mdlIndent.IndentRound"
Public Const COMMAND_TITLE1 = "Drafting.mdlStyle.SetTitle1"
Public Const COMMAND_TITLE2 = "Drafting.mdlStyle.SetTitle2"
Public Const COMMAND_TITLE3 = "Drafting.mdlStyle.SetTitle3"
Public Const COMMAND_TITLE4 = "Drafting.mdlStyle.SetTitle4"
Public Const COMMAND_TITLE5 = "Drafting.mdlStyle.SetTitle5"
Public Const COMMAND_BODY1 = "Drafting.mdlStyle.SetBody1"
Public Const COMMAND_BODY2 = "Drafting.mdlStyle.SetBody2"
Public Const COMMAND_BODY3 = "Drafting.mdlStyle.SetBody3"
Public Const COMMAND_BODY4 = "Drafting.mdlStyle.SetBody4"
Public Const COMMAND_BODY5 = "Drafting.mdlStyle.SetBody5"
Public Const COMMAND_CLEAR_STYLE = "Drafting.mdlStyle.ClearStyle"

Public Enum RECOMMEND
    KEY_RIGHT = wdKeyControl + wdKeyShift + wdKeyK
    KEY_LEFT = wdKeyControl + wdKeyShift + wdKeyJ
    KEY_PRIMARY_RIGHT = wdKeyControl + wdKeyShift + wdKeyR
    KEY_PRIMARY_LEFT = wdKeyControl + wdKeyShift + wdKeyE
    KEY_SECONDARY_RIGHT = wdKeyControl + wdKeyShift + wdKeyF
    KEY_SECONDARY_LEFT = wdKeyControl + wdKeyShift + wdKeyD
    KEY_ROUND = wdKeyControl + wdKeyShift + wdKeyI
    KEY_TITLE1 = wdKeyAlt + wdKey1
    KEY_TITLE2 = wdKeyAlt + wdKey2
    KEY_TITLE3 = wdKeyAlt + wdKey3
    KEY_TITLE4 = wdKeyAlt + wdKey4
    KEY_TITLE5 = wdKeyAlt + wdKey5
    KEY_BODY1 = wdKeyAlt + wdKey6
    KEY_BODY2 = wdKeyAlt + wdKey7
    KEY_BODY3 = wdKeyAlt + wdKey8
    KEY_BODY4 = wdKeyAlt + wdKey9
    KEY_BODY5 = wdKeyAlt + wdKey0
    KEY_CLEAR_STYLE = wdKeyControl + wdKeyCloseSquareBrace
End Enum

Public styleNames(12) As String
Public chkGothics(12) As MSForms.CheckBox
Public chkBolds(12) As MSForms.CheckBox
