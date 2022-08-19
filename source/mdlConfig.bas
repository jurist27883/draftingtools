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
Public Const TAG_INDENT = "indent"

Public Const CSIDL_FONTS = 20

Public Const COMMAND_NAME_RIGHT = "全体インデント 右"
Public Const COMMAND_NAME_LEFT = "全体インデント 左"
Public Const COMMAND_NAME_PRIMARY_RIGHT = "1行目インデント 右"
Public Const COMMAND_NAME_PRIMARY_LEFT = "1行目インデント 左"
Public Const COMMAND_NAME_SECONDARY_RIGHT = "2行目以下インデント 右"
Public Const COMMAND_NAME_SECONDARY_LEFT = "2行目以下インデント 左"
Public Const COMMAND_NAME_ROUND = "インデント値丸め"
Public Const COMMAND_NAME_TITLE1 = "見出し 第1"
Public Const COMMAND_NAME_TITLE2 = "見出し 1"
Public Const COMMAND_NAME_TITLE3 = "見出し (1)"
Public Const COMMAND_NAME_TITLE4 = "見出し ア"
Public Const COMMAND_NAME_TITLE5 = "見出し (ア)"
Public Const COMMAND_NAME_BODY1 = "本文  第1"
Public Const COMMAND_NAME_BODY2 = "本文  1"
Public Const COMMAND_NAME_BODY3 = "本文  (1)"
Public Const COMMAND_NAME_BODY4 = "本文  ア"
Public Const COMMAND_NAME_BODY5 = "本文  (ア)"
Public Const COMMAND_NAME_CLEAR_STYLE = "書式クリア"

Public Const COMMAND_RIGHT = "DraftingTools.mdlIndent.IndentRight"
Public Const COMMAND_LEFT = "DraftingTools.mdlIndent.IndentLeft"
Public Const COMMAND_PRIMARY_RIGHT = "DraftingTools.mdlIndent.IndentPrimaryRight"
Public Const COMMAND_PRIMARY_LEFT = "DraftingTools.mdlIndent.IndentPrimaryLeft"
Public Const COMMAND_SECONDARY_RIGHT = "DraftingTools.mdlIndent.IndentSecondaryRight"
Public Const COMMAND_SECONDARY_LEFT = "DraftingTools.mdlIndent.IndentSecondaryLeft"
Public Const COMMAND_ROUND = "DraftingTools.mdlIndent.IndentRound"
Public Const COMMAND_TITLE1 = "DraftingTools.mdlStyle.SetTitle1"
Public Const COMMAND_TITLE2 = "DraftingTools.mdlStyle.SetTitle2"
Public Const COMMAND_TITLE3 = "DraftingTools.mdlStyle.SetTitle3"
Public Const COMMAND_TITLE4 = "DraftingTools.mdlStyle.SetTitle4"
Public Const COMMAND_TITLE5 = "DraftingTools.mdlStyle.SetTitle5"
Public Const COMMAND_BODY1 = "DraftingTools.mdlStyle.SetBody1"
Public Const COMMAND_BODY2 = "DraftingTools.mdlStyle.SetBody2"
Public Const COMMAND_BODY3 = "DraftingTools.mdlStyle.SetBody3"
Public Const COMMAND_BODY4 = "DraftingTools.mdlStyle.SetBody4"
Public Const COMMAND_BODY5 = "DraftingTools.mdlStyle.SetBody5"
Public Const COMMAND_CLEAR_STYLE = "DraftingTools.mdlStyle.ClearStyle"

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

Public styleNames(9) As String
Public chkGothics(9) As MSForms.CheckBox
Public chkBolds(9) As MSForms.CheckBox
Public txtIndents(4) As MSForms.TextBox

