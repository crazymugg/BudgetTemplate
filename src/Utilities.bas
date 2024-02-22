Attribute VB_Name = "Utilities"
Option Explicit

Public Sub ShowAdmin()
    Dim AdminArea As Range

    Set AdminArea = ActiveSheet.Columns("E:L")

    If AdminArea.Hidden = False Then
        AdminArea.Hidden = True
        ActiveSheet.Buttons("AdminButton").Text = "Show Admin"
    Else
        AdminArea.Hidden = False
        ActiveSheet.Buttons("AdminButton").Text = "Hide Admin"
    End If

End Sub

