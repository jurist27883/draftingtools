Attribute VB_Name = "mdlFunction"
'@Folder "other"
Option Explicit

Function isKatakana(ByVal s As String) As Boolean
    Select Case AscW(s)
    Case -143 To -99, 12450 To 12531
        isKatakana = True
    Case Else
        isKatakana = False
    End Select
End Function

Function isSpace(ByVal s As String) As Boolean
    If s = " " Or s = "Å@" Then
        isSpace = True
    Else
        isSpace = False
    End If
End Function

Sub SelectText(targetTextBox As textBox)
    With targetTextBox
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Property Get SelectedList(targetListBox As ListBox, column As Long) As Variant
    If targetListBox.ListIndex = -1 Then
        SelectedList = Nothing
    Else
        SelectedList = targetListBox.List(targetListBox.ListIndex, column)
    End If
End Property

Property Set SelectedList(targetListBox As ListBox, column As Long, value As Variant)
    If targetListBox.ListIndex Is Nothing Then
        Exit Sub
    Else
        targetListBox.List(targetListBox.ListIndex, column) = value
    End If
End Property

