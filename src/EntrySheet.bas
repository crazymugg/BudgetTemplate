Attribute VB_Name = "EntrySheet"
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
        .Range("A1", "A1").Value = "Date"
        .Range("B1", "B1").Value = "Cost"
        .Range("C1", "C1").Value = "Place"
        .Columns("C").ColumnWidth = 40
        .Range("D1", "D1").Value = "Location"
        .Columns("D").ColumnWidth = 25
        .Range("E1", "E1").Value = "Method"
        .Range("F1", "F1").Value = "Notes"
        .Columns("F").ColumnWidth = 50
        
        .ListObjects.Add(xlSrcRange, Range("A1:F1"), , xlYes).Name = "Table" + Year
    End With
    
    
End Sub


Public Sub AddRowTable(ByVal Year As String, ByVal EntryData As Entry)
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
