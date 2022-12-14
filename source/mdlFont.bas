Attribute VB_Name = "mdlFont"
'@Folder "main.style"
Option Explicit

Public Const MSMIN = "lr ¾©"
Public Const MSGOTHIC = "lr SVbN"
Public Const YUMIN = "à¾©"
Public Const YUGOTHIC = "àSVbN"
Public Const BIZUDMIN = "BIZ UD¾© Medium"
Public Const BIZUDGOTHIC = "BIZ UDSVbN"

Property Get CounterFontName(fontName As String) As String
    '¾©ESVbNØÖ¦
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

