Attribute VB_Name = "Accounts"
Option Explicit

Public Sub CreateAccount()
    MsgBox ("Placeholder Text")
End Sub

Public Sub EditAccount()
    MsgBox ("Placeholder Text")
End Sub

Public Sub DeleteAccount()
    MsgBox ("Placeholder Text")
End Sub

Public Sub TestAccount()
    Dim MyTable As ListObject
    Dim LoopColumn As ListColumn
    
    Set MyTable = ActiveSheet.ListObjects("AccountsTable")
    
    For Each LoopColumn In MyTable.ListColumns
    
        Debug.Print LoopColumn.Name
        Debug.Print LoopColumn.Index
        Debug.Print LoopColumn.Parent
    
    Next
    
    Set LoopColumn = MyTable.ListColumns.Add
    LoopColumn.Name = "TestAccount4"
    LoopColumn.Range(0, 1).Value = "Test4"
    
End Sub


Public Sub Crap()
    Dim Poop As test
    Set Poop = New test
    With Poop
        .Title = "Poop"
        .Index = 1
        .Length = 10
        .Category = "Credit"
    End With
    
    
    Debug.Print (Poop.Title)
End Sub
