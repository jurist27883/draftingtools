Attribute VB_Name = "mdlListConverter"
'@Folder("main")
Option Explicit

Sub ConvertListNumbers()
    '箇条書きの段落番号等をテキストに変換
    ActiveDocument.ConvertNumbersToText wdNumberParagraph
    With Application.Options
        '入力オートフォーマット/入力中に自動で変更する項目/行の始まりのスペースを字下げに変更する
        .AutoFormatApplyFirstIndents = True
        '入力オートフォーマット/入力中に自動で書式設定する項目/箇条書き（行頭文字）
        .AutoFormatAsYouTypeApplyBulletedLists = False
        '入力オートフォーマット/入力中に自動で書式設定する項目/箇条書き（段落番号）
        .AutoFormatAsYouTypeApplyNumberedLists = False
        'オートフォーマット/自動で適切なスタイルを設定する箇所/箇条書き（行頭文字）
        .AutoFormatApplyLists = False
        'オートフォーマット/自動で適切なスタイルを設定する箇所/リストのスタイル
        .AutoFormatApplyBulletedLists = False
    End With
    
End Sub
