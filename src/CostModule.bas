Attribute VB_Name = "CostModule"
Option Explicit

Public Sub TestCost()
    Dim NewCost As New CostEntry
    
    With NewCost:
        .EntryDate = Date
        .EntryCost = 27.11
        .EntryPlace = "Shell"
        .EntryLocation = "Hendersonville, TN"
        .EntryMethod = "SW"
        .EntryNotes = ""
    End With
    
    Debug.Print (NewCost.EntryPlace)
    Debug.Print (VarType(NewCost.EntryCost))

End Sub


Public Sub ShowCostEntryForm()
    CostForm.Show
    Debug.Print ("Help")
End Sub

