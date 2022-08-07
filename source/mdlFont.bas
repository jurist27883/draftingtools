Attribute VB_Name = "mdlFont"
'@Folder "main"
Option Explicit

Public Const MSMIN = "�l�r ����"
Public Const MSGOTHIC = "�l�r �S�V�b�N"
Public Const YUMIN = "������"
Public Const YUGOTHIC = "���S�V�b�N"
Public Const BIZUDMIN = "BIZ UD���� Medium"
Public Const BIZUDGOTHIC = "BIZ UD�S�V�b�N"

Property Get CounterFontName(fontName As String) As String
    '�����E�S�V�b�N�ؑւ�
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



