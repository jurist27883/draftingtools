Attribute VB_Name = "mdlFunction"
'@Folder "other"
Option Explicit
Sub PushArray(ByRef arr As Variant, pushingData As Variant)

    If IsArray(arr) = False Then
        Exit Sub
    End If
    
    Dim length As Long
    length = UBound(arr) + 1
    ReDim Preserve arr(length)
    
    Dim i As Long
    If IsObject(arr(0)) Then
        Set arr(length) = pushingData
    Else
        arr(length) = pushingData
    End If
    
End Sub

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

