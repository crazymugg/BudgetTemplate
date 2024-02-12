Attribute VB_Name = "Button"
Option Explicit

Public Sub NewEntryButton()
    Call Placeholder
End Sub


Public Sub NewYearButton()
    Call EntrySheet.NewYearSheet
End Sub


Public Sub Placeholder()
    Dim Response As String
    Response = InputBox("Help", "This is the title?", "return")
    Debug.Print (Response)
    'MsgBox ("Placholder text here!")
End Sub


Public Sub NewAccountButton()
    Call Placeholder
End Sub


