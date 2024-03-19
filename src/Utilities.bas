Attribute VB_Name = "Utilities"
Option Explicit

Public Sub ShowAdmin()
    Dim AdminCols As Range
    Dim AdminRows As Range

    Set AdminCols = ActiveSheet.Columns("I:S")
    Set AdminRows = ActiveSheet.Rows("23:29")

    If AdminCols.Hidden = False Then
        AdminCols.Hidden = True
        AdminRows.Hidden = True
        ActiveSheet.Buttons("AdminButton").Text = "Show Admin"
    Else
        AdminCols.Hidden = False
        AdminRows.Hidden = False
        ActiveSheet.Buttons("AdminButton").Text = "Hide Admin"
    End If

End Sub


Public Sub ShowHelp()
    Dim HelpArea As Range

    Set HelpArea = ActiveSheet.Columns("A:D")

    If HelpArea.Hidden = False Then
        HelpArea.Hidden = True
        ActiveSheet.Buttons("HelpButton").Text = "Show Help"
    Else
        HelpArea.Hidden = False
        ActiveSheet.Buttons("HelpButton").Text = "Hide Help"
    End If

End Sub

