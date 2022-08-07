Attribute VB_Name = "mdlRibbon"
'@Folder "ribbon"
Option Explicit

Public rb As IRibbonUI

Public Const TABNAME = "drafting"

Sub Ribbon_OnLoad(Ribbon As IRibbonUI)
    'リボン初期設定
    Set rb = Ribbon
End Sub

Public Sub ActivateTab()
    On Error Resume Next    '初回起動時にエラーとなることを回避するため
    mdlRibbon.rb.ActivateTab (mdlRibbon.TABNAME)
End Sub

Sub IndentFirstlineRight_onAction(control As IRibbonControl)
    IndentFirstLineRight
End Sub

Sub IndentFirstlineLeft_onAction(control As IRibbonControl)
    IndentFirstLineLeft
End Sub

Sub IndentPrimaryRight_onAction(control As IRibbonControl)
    IndentPrimaryRight
End Sub

Sub IndentPrimaryLeft_onAction(control As IRibbonControl)
    IndentPrimaryLeft
End Sub

Sub IndentSecondaryRight_onAction(control As IRibbonControl)
    IndentSecondaryRight
End Sub

Sub IndentSecondaryLeft_onAction(control As IRibbonControl)
    IndentSecondaryLeft
End Sub

Sub IndentRight_onAction(control As IRibbonControl)
    IndentRight
End Sub

Sub IndentLeft_onAction(control As IRibbonControl)
    IndentLeft
End Sub


Sub IndentRound_onAction(control As IRibbonControl)
    IndentRound
End Sub

Sub Title1_onAction(control As IRibbonControl)
    mdlStyle.SetTitle1
End Sub
Sub Title2_onAction(control As IRibbonControl)
    mdlStyle.SetTitle2
End Sub
Sub Title3_onAction(control As IRibbonControl)
    mdlStyle.SetTitle3
End Sub
Sub Title4_onAction(control As IRibbonControl)
    mdlStyle.SetTitle4
End Sub
Sub Title5_onAction(control As IRibbonControl)
    mdlStyle.SetTitle5
End Sub
Sub Body1_onAction(control As IRibbonControl)
    mdlStyle.SetBody1
End Sub
Sub Body2_onAction(control As IRibbonControl)
    mdlStyle.SetBody2
End Sub
Sub Body3_onAction(control As IRibbonControl)
    mdlStyle.SetBody3
End Sub
Sub Body4_onAction(control As IRibbonControl)
    mdlStyle.SetBody4
End Sub
Sub Body5_onAction(control As IRibbonControl)
    mdlStyle.SetBody5
End Sub
Sub ClearStyle_onAction(control As IRibbonControl)
    mdlStyle.ClearStyle
End Sub

Sub Config_onAction(control As IRibbonControl)
    frmConfig.Show
End Sub
