VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfig 
   Caption         =   "設定"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9690
   OleObjectBlob   =   "frmConfig.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "config.form"
Option Explicit

Private Sub SetStyleFonts(styleName As String, dstGothic As Boolean, dstBold As Boolean)
    'ActiveDocumentのStyleのフォント設定
    'bug activedocumentに反映されない
    If dstGothic <> mdlStyle.IsGothic(styleName) Then
        mdlStyle.ToggleFontFamily styleName
    End If
    ActiveDocument.Styles(styleName).Font.Bold = dstBold
End Sub

Private Sub SetStyleIndents(styleName As String, dstlIndentCount As Long)
    'ActiveDocumentのStyleのインデント設定
    Select Case styleName
    Case TITLE1, TITLE2, TITLE3, TITLE4, TITLE5
        mdlStyle.IndentCount(styleName) = dstlIndentCount
    Case BODY1
        mdlStyle.IndentCount(styleName) = dstlIndentCount + 2
    Case BODY2, BODY3, BODY4, BODY5
        mdlStyle.IndentCount(styleName) = dstlIndentCount + 1
    End Select
End Sub

Private Sub SaveStyleXml(styleName As String, dstGothic As Boolean, dstBold As Boolean, dstlIndentCount)
    'xmlファイル削除・登録
    #If RELEASE Then
        Dim dicElements As Object
        Set dicElements = CreateObject("Scripting.Dictionary")
    #Else
        Dim dicElements As Scripting.Dictionary
        Set dicElements = New Scripting.Dictionary
    #End If
    
    mdlXml.RemoveNode TAG_STYLE, TAG_STYLE_NAME, styleName
    
    dicElements(TAG_STYLE_NAME) = styleName
    dicElements(TAG_GOTHIC) = dstGothic
    dicElements(TAG_BOLD) = dstBold
    Select Case styleName
    Case TITLE1, TITLE2, TITLE3, TITLE4, TITLE5
        dicElements(TAG_INDENT) = dstlIndentCount
    Case BODY1
        dicElements(TAG_INDENT) = dstlIndentCount + 2
    Case BODY2, BODY3, BODY4, BODY5
        dicElements(TAG_INDENT) = dstlIndentCount + 1
    End Select
    mdlXml.AddNode TAG_STYLE, dicElements
End Sub

Private Sub cmdCancel_Click()
    Unload frmConfig
End Sub

Private Sub lstClass_Change()
    '分類リスト行選択時のキーリスト絞込み
    If lstClass.ListIndex <> -1 Then
        Select Case lstClass.ListIndex
        Case CLASS_INDEX.INSDENT_NUMBER
            lstKeys.ListIndex = 0
        Case CLASS_INDEX.STYLE_NUMBER
            lstKeys.ListIndex = 7
        End Select
        cmdRecommendKey.Enabled = True
    End If

End Sub

Private Sub lstKeys_Change()
    'キーリスト行選択時のボタン処理
    If lstKeys.ListIndex = -1 Then
        cmdRecommendKey.Enabled = False
        cmdResisterKey.Enabled = False
        cmdResetKey.Enabled = False
    Else
        cmdRecommendKey.Enabled = True
        If mdlListBoxKeys.SelectedItem(CL.BINDING_KEY_CODE) = 0 Then
            cmdResetKey.Enabled = False
        Else
            cmdResetKey.Enabled = True
        End If
    End If
End Sub

Private Sub mpConfig_Change()
    'タブ切替時
    Select Case mpConfig.SelectedItem.Name
    Case "pgKey"
        mdlListBoxKeys.Initialize
    Case "pgStyle"
        InitializeStyleTab
    End Select
End Sub

Private Sub InitializeForm()
    'フォーム呼出時初期設定
    styleNames(0) = TITLE1
    styleNames(1) = TITLE2
    styleNames(2) = TITLE3
    styleNames(3) = TITLE4
    styleNames(4) = TITLE5
    styleNames(5) = BODY1
    styleNames(6) = BODY2
    styleNames(7) = BODY3
    styleNames(8) = BODY4
    styleNames(9) = BODY5
    
    Set chkGothics(0) = frmConfig.chkGothicTitle1
    Set chkGothics(1) = frmConfig.chkGothicTitle2
    Set chkGothics(2) = frmConfig.chkGothicTitle3
    Set chkGothics(3) = frmConfig.chkGothicTitle4
    Set chkGothics(4) = frmConfig.chkGothicTitle5
    Set chkGothics(5) = frmConfig.chkGothicBody1
    Set chkGothics(6) = frmConfig.chkGothicBody2
    Set chkGothics(7) = frmConfig.chkGothicBody3
    Set chkGothics(8) = frmConfig.chkGothicBody4
    Set chkGothics(9) = frmConfig.chkGothicBody5
    
    Set chkBolds(0) = frmConfig.chkBoldTitle1
    Set chkBolds(1) = frmConfig.chkBoldTitle2
    Set chkBolds(2) = frmConfig.chkBoldTitle3
    Set chkBolds(3) = frmConfig.chkBoldTitle4
    Set chkBolds(4) = frmConfig.chkBoldTitle5
    Set chkBolds(5) = frmConfig.chkBoldBody1
    Set chkBolds(6) = frmConfig.chkBoldBody2
    Set chkBolds(7) = frmConfig.chkBoldBody3
    Set chkBolds(8) = frmConfig.chkBoldBody4
    Set chkBolds(9) = frmConfig.chkBoldBody5
    
    Set txtIndents(0) = frmConfig.txtIndentTitle1
    Set txtIndents(1) = frmConfig.txtIndentTitle2
    Set txtIndents(2) = frmConfig.txtIndentTitle3
    Set txtIndents(3) = frmConfig.txtIndentTitle4
    Set txtIndents(4) = frmConfig.txtIndentTitle5
    
End Sub

Private Sub InitializeStyleTab()
    'スタイルタブ初期化
    Dim i As Long
    For i = 0 To 9
        mdlStyle.ResisterStyle styleNames(i)
        
        If mdlStyle.IsGothic(styleNames(i)) Then
            chkGothics(i).value = True
        Else
            chkGothics(i).value = False
        End If
        
        If mdlStyle.IsBold(styleNames(i)) Then
            chkBolds(i).value = True
        Else
            chkBolds(i).value = False
        End If
    Next
    
    For i = 0 To UBound(txtIndents)
        txtIndents(i).value = mdlStyle.IndentCount(styleNames(i))
    Next
End Sub

Private Sub cmdSetStyleSequence_Click()
    'スタイルをactivedocumentに設定・xml登録
    Dim i As Long
    For i = 0 To UBound(styleNames)
        SetStyleFonts styleNames(i), chkGothics(i).value, chkBolds(i).value
        SetStyleIndents styleNames(i), txtIndents(i Mod 5).value
        SaveStyleXml styleNames(i), chkGothics(i).value, chkBolds(i).value, txtIndents(i Mod 5).value
    Next
End Sub

Private Sub cmdCancelStyle_Click()
    Unload frmConfig
End Sub

Private Sub SetControls()
    'フォーム上のコントロールへの反映
    If lstKeys.ListIndex = -1 Then
        cmdRecommendKey.Enabled = False
        cmdResisterKey.Enabled = False
        cmdResetKey.Enabled = False
        Exit Sub
    End If

    Dim bindingKeyCode As Long
    bindingKeyCode = mdlKey.BindingCode(mdlListBoxKeys.SelectedItem(COMMAND_STRING))

    If bindingKeyCode = 0 Then
        mdlListBoxKeys.SelectedItem(CL.BINDING_KEY_CODE) = 0
        mdlListBoxKeys.SelectedItem(CL.BINDING_KEY_STRING) = ""
        cmdResetKey.Enabled = False
    Else
        mdlListBoxKeys.SelectedItem(CL.BINDING_KEY_CODE) = bindingKeyCode
        mdlListBoxKeys.SelectedItem(CL.BINDING_KEY_STRING) = KeyString(bindingKeyCode)
        cmdResetKey.Enabled = True
    End If
        
    lblPresentCommand = ""
    txtAssigningKey = ""
End Sub

Private Sub txtAssigningKey_Change()
    '文字のみ/Shift+文字 のキー押下げがなされた場合にkeydownイベント内で
    'テキストボックスのtextプロパティに空文字を設定しても反映しないことに対する処理と
    'ボタンの処理
    If assigningKeyCode = 0 Then
        txtAssigningKey.text = ""
        lblPresentCommand = ""
        cmdResisterKey.Enabled = False
    Else
        cmdResisterKey.Enabled = True
    End If
End Sub

Private Sub txtAssigningKey_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'ショートカットキー候補押下げ時
    If ((Shift And MASK.CTRLS) Or (Shift And MASK.ALTS)) = 0 Then
        assigningKeyCode = 0
        Exit Sub
    End If
    
    If Shift And MASK.SHIFTS Then
        KeyCode = KeyCode + wdKeyShift
    End If
    If Shift And MASK.CTRLS Then
        KeyCode = KeyCode + wdKeyControl
    End If
    If Shift And MASK.ALTS Then
        KeyCode = KeyCode + wdKeyAlt
    End If
    
    #If RELEASE Then
        CustomizationContext = NormalTemplate
    #Else
        CustomizationContext = ThisDocument
    #End If
    
    Dim presentCommand As String
    assigningKeyCode = KeyCode
    presentCommand = mdlKey.BindingCommand(CLng(KeyCode))
    If presentCommand <> "" Then
        lblPresentCommand = "現在の割当て：" + presentCommand
    End If
    txtAssigningKey.text = KeyString(KeyCode)

End Sub

Private Sub cmdResisterKey_Click()
    'キー設定
    If assigningKeyCode = 0 Then
        MsgBox "登録するショートカットキーをテキストボックスに入力してください。"
        Exit Sub
    End If
    '既登録分の解除
    mdlKey.ResetKey mdlListBoxKeys.SelectedItem(CL.COMMAND_STRING), mdlListBoxKeys.SelectedItem(CL.BINDING_KEY_CODE)
    'configファイル登録
    #If RELEASE Then
        Dim dicElements As Object
        Set dicElements = CreateObject("Scripting.Dictionary")
    #Else
        Dim dicElements As Scripting.Dictionary
        Set dicElements = New Scripting.Dictionary
    #End If
    dicElements(TAG_KEYCODE) = assigningKeyCode
    dicElements(TAG_FORMER_COMMAND) = mdlKey.BindingCommand(assigningKeyCode)
    mdlXml.AddNode TAG_KEYBINDING, dicElements
    '新規キーバインディング登録
    mdlKey.BindingCode(mdlListBoxKeys.SelectedItem(CL.COMMAND_STRING)) = assigningKeyCode
    'フォーム更新
    SetControls
    
    assigningKeyCode = 0
End Sub

Private Sub cmdResetKey_Click()
    'キー設定解除
    
    If lstKeys.ListIndex = -1 Then
        MsgBox "リストから解除するコマンドを選択してください。"
        Exit Sub
    End If
    If mdlListBoxKeys.SelectedItem(CL.BINDING_KEY_CODE) = 0 Then
        Exit Sub
    End If
    '解除
    mdlKey.ResetKey mdlListBoxKeys.SelectedItem(CL.COMMAND_STRING), mdlListBoxKeys.SelectedItem(CL.BINDING_KEY_CODE)
    'フォーム更新
    SetControls
    
    assigningKeyCode = 0
End Sub

Private Sub cmdRecommendKey_Click()
    'キー推奨
    If lstKeys.ListIndex = -1 Then
        Exit Sub
    End If
    assigningKeyCode = mdlListBoxKeys.SelectedItem(CL.RECOMMEND_KEY_CODE)
    txtAssigningKey = KeyString(assigningKeyCode)
End Sub

Private Sub txtAssigningKey_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '文字列選択
    mdlFunction.SelectText txtAssigningKey
End Sub

Private Sub CheckIndentTitles(targetTextbox As MSForms.TextBox, Cancel As MSForms.ReturnBoolean)
    If IsNumeric(targetTextbox.text) Then
        If CLng(targetTextbox.text) >= 0 And CLng(targetTextbox.text) <= 10 _
        And Int(targetTextbox.text) = targetTextbox.text Then
            targetTextbox = StrConv(targetTextbox.text, vbNarrow)
        Else
            MsgBox "0から10までの整数を入力してください。"
            Cancel = True
            mdlFunction.SelectText targetTextbox
        End If
    Else
        MsgBox "半角数字を入力してください。"
        Cancel = True
        mdlFunction.SelectText targetTextbox
    End If
End Sub

Private Sub txtIndentTitle1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    CheckIndentTitles txtIndentTitle1, Cancel
End Sub
Private Sub txtIndentTitle2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    CheckIndentTitles txtIndentTitle2, Cancel
End Sub
Private Sub txtIndentTitle3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    CheckIndentTitles txtIndentTitle3, Cancel
End Sub
Private Sub txtIndentTitle4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    CheckIndentTitles txtIndentTitle4, Cancel
End Sub
Private Sub txtIndentTitle5_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    CheckIndentTitles txtIndentTitle5, Cancel
End Sub
Private Sub txtIndentTitle1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
     mdlFunction.SelectText txtIndentTitle1
End Sub
Private Sub txtIndentTitle2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
     mdlFunction.SelectText txtIndentTitle2
End Sub
Private Sub txtIndentTitle3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
     mdlFunction.SelectText txtIndentTitle3
End Sub
Private Sub txtIndentTitle4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
     mdlFunction.SelectText txtIndentTitle4
End Sub
Private Sub txtIndentTitle5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
     mdlFunction.SelectText txtIndentTitle5
End Sub

Private Sub UserForm_Initialize()
    'フォーム呼出時
    InitializeForm
    InitializeStyleTab
    
    mdlListBoxClass.Initialize
    mdlListBoxKeys.Initialize
End Sub

