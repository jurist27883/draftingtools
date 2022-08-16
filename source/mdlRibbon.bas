Attribute VB_Name = "mdlRibbon"
'@Folder "ribbon"
Option Explicit

Private rb As IRibbonUI
Private tgTrackingFormatPressed As Boolean

Public Const TABNAME = "drafting"

Private Sub Ribbon_OnLoad(Ribbon As IRibbonUI)
    'ÉäÉ{Éìèâä˙ê›íË
    Set rb = Ribbon
    tgTrackingFormatPressed = ActiveDocument.TrackFormatting
    rb.Invalidate
End Sub
Private Sub ToggleTrackingFormat_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = tgTrackingFormatPressed
End Sub
Private Sub ToggleTrackingFormat_onAction(control As IRibbonControl, pressed As Boolean)
    mdlRevision.ToggleTrackingFormat
    tgTrackingFormatPressed = pressed
End Sub
Private Sub AcceptChangedFormat_onAction(control As IRibbonControl)
    mdlRevision.AcceptChangedFormat
End Sub

Private Sub AcceptRevisions_onAction(control As IRibbonControl)
    mdlRevision.AcceptRevisions
End Sub

Private Sub ConvertListNumbers_onAction(control As IRibbonControl)
    mdlListConverter.ConvertListNumbers
End Sub
Private Sub SetStyles_onAction(control As IRibbonControl)
    mdlStyleSetter.SetStyles
End Sub
Private Sub ChangeDigit_onAction(control As IRibbonControl)
    mdlDigit.ChangeDigit
End Sub
Private Sub CheckVersion_onAction(control As IRibbonControl)
    mdlVersionChecker.CheckUpdate
End Sub
Private Sub IndentFirstlineRight_onAction(control As IRibbonControl)
    IndentFirstLineRight
End Sub
Private Sub IndentFirstlineLeft_onAction(control As IRibbonControl)
    IndentFirstLineLeft
End Sub
Private Sub IndentPrimaryRight_onAction(control As IRibbonControl)
    IndentPrimaryRight
End Sub
Private Sub IndentPrimaryLeft_onAction(control As IRibbonControl)
    IndentPrimaryLeft
End Sub
Private Sub IndentSecondaryRight_onAction(control As IRibbonControl)
    IndentSecondaryRight
End Sub
Private Sub IndentSecondaryLeft_onAction(control As IRibbonControl)
    IndentSecondaryLeft
End Sub
Private Sub IndentRight_onAction(control As IRibbonControl)
    IndentRight
End Sub
Private Sub IndentLeft_onAction(control As IRibbonControl)
    IndentLeft
End Sub
Private Sub IndentRound_onAction(control As IRibbonControl)
    IndentRound
End Sub
Private Sub Title1_onAction(control As IRibbonControl)
    mdlStyle.SetTitle1
End Sub
Private Sub Title2_onAction(control As IRibbonControl)
    mdlStyle.SetTitle2
End Sub
Private Sub Title3_onAction(control As IRibbonControl)
    mdlStyle.SetTitle3
End Sub
Private Sub Title4_onAction(control As IRibbonControl)
    mdlStyle.SetTitle4
End Sub
Private Sub Title5_onAction(control As IRibbonControl)
    mdlStyle.SetTitle5
End Sub
Private Sub Body1_onAction(control As IRibbonControl)
    mdlStyle.SetBody1
End Sub
Private Sub Body2_onAction(control As IRibbonControl)
    mdlStyle.SetBody2
End Sub
Private Sub Body3_onAction(control As IRibbonControl)
    mdlStyle.SetBody3
End Sub
Private Sub Body4_onAction(control As IRibbonControl)
    mdlStyle.SetBody4
End Sub
Private Sub Body5_onAction(control As IRibbonControl)
    mdlStyle.SetBody5
End Sub
Private Sub ClearStyle_onAction(control As IRibbonControl)
    mdlStyle.ClearStyle
End Sub
Private Sub Config_onAction(control As IRibbonControl)
    frmConfig.Show
End Sub
