Attribute VB_Name = "OLD"
Option Explicit

Public Sub NewYearSheet()
    Dim LastSheetName As String
    Dim NewWS As Worksheet
    Dim CalcYear As Integer
    Dim Year As String
    
    LastSheetName = Sheets(Sheets.Count).Name
    
    CalcYear = CInt(LastSheetName) + 1
    
    Year = CStr(CalcYear)

    Worksheets.Add(After:=Sheets(Sheets.Count)).Name = Year

    Set NewWS = Sheets(Year)
    
    With NewWS
        .Range("A1", "A1").Value = "ID"
        .Range("B1", "B1").Value = "Date"
        .Range("C1", "C1").Value = "Cost"
        .Range("D1", "D1").Value = "Place"
        .Columns("D").ColumnWidth = 40
        .Range("E1", "E1").Value = "Location"
        .Columns("E").ColumnWidth = 25
        .Range("F1", "F1").Value = "Method"
        .Range("G1", "G1").Value = "Notes"
        .Columns("G").ColumnWidth = 50
        
        .ListObjects.Add(xlSrcRange, Range("A1:G1"), , xlYes).Name = "Table" + Year
    End With
    
    
End Sub


Public Sub AddRowTable(ByVal Year As String, ByVal EntryData As Integer)
    Dim WS As Worksheet
    Dim TblName As String
    Dim Tbl As ListObject
    Dim TblRow As ListRow
    
    Set WS = Sheets(Year)
    Set TblName = "Table" + Year
    Set Tbl = WS.ListObjects(Tbl_Name)
    Set TblRow = Tbl.ListRows.Add()

    TblRow.Range(1, 1).Value = "Test A"
    TblRow.Range(1, 2).Value = "Test B"
    TblRow.Range(1, 3).Value = "Test C"
    TblRow.Range(1, 4).Value = "Test D"
End Sub


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



Public Sub CreateAccount()
    'MsgBox ("Placeholder Text")
    AddAccount.Show
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
    Dim Poop As Test
    Set Poop = New Test
    With Poop
        .Title = "Poop"
        .Index = 1
        .Length = 10
        .Category = "Credit"
    End With
    
    
    Debug.Print (Poop.Title)
End Sub



Public Sub DateFormatting()
    Dim MyDate As Date
    Dim MyDateString As String
    Dim DateArray() As String
    
    MyDate = Date
    MyDateString = CStr(MyDate)
    
    DateArray = Split(MyDateString, "-")
    
    Debug.Print (DateArray(2))
    
    
    
End Sub








