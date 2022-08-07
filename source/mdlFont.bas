Attribute VB_Name = "mdlFont"
'@Folder "main"
Option Explicit

Public Const MSMIN = "ＭＳ 明朝"
Public Const MSGOTHIC = "ＭＳ ゴシック"
Public Const YUMIN = "游明朝"
Public Const YUGOTHIC = "游ゴシック"
Public Const BIZUDMIN = "BIZ UD明朝 Medium"
Public Const BIZUDGOTHIC = "BIZ UDゴシック"

Property Get CounterFontName(fontName As String) As String
    '明朝・ゴシック切替え
    Select Case fontName
    Case MSMIN
        CounterFontName = MSGOTHIC
    Case MSGOTHIC
        CounterFontName = MSMIN
    Case YUMIN
        CounterFontName = YUGOTHIC
    Case YUGOTHIC
        CounterFontName = YUMIN
    Case BIZUDMIN
        CounterFontName = BIZUDGOTHIC
    Case BIZUDGOTHIC
        CounterFontName = BIZUDMIN
    Case Else
        CounterFontName = MSMIN
    End Select
End Property



