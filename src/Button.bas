Attribute VB_Name = "Button"
Option Explicit


Public Sub Placeholder()
    Dim Response As String
    Response = InputBox("Help", "This is the title?", "return")
    Debug.Print (Response)
    'MsgBox ("Placholder text here!")
End Sub
