VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} YearForm 
   Caption         =   "Add a New Year"
   ClientHeight    =   600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6810
   OleObjectBlob   =   "YearForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "YearForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SubmitButton_Click()
    Application.ScreenUpdating = False

    'Create Table
    Dim YearName As String
    Dim NewWS As Worksheet
    
    'Add New Sheet titled the year
    YearName = YearForm.YearBox.Value
    Worksheets.Add(After:=Sheets(Sheets.Count)).Name = YearName
    Set NewWS = Sheets(YearName)
    
    'Add Header Row
    With NewWS
        .Range("A1", "A1").Value = "ID"
        .Columns("A").ColumnWidth = 6
        
        .Range("B1", "B1").Value = "Date"
        .Columns("B").ColumnWidth = 10
        .Columns("B").NumberFormat = "m/d/yy;@"
        
        .Range("C1", "C1").Value = "Cost"
        .Columns("C").ColumnWidth = 10
        .Columns("C").NumberFormat = "#,##0.00"
        
        .Range("D1", "D1").Value = "Place"
        .Columns("D").ColumnWidth = 30
        
        .Range("E1", "E1").Value = "Location"
        .Columns("E").ColumnWidth = 25
        
        .Range("F1", "F1").Value = "Category"
        .Columns("F").ColumnWidth = 20
        
        .Range("G1", "G1").Value = "Method"
        
        .Range("H1", "H1").Value = "Notes"
        .Columns("H").ColumnWidth = 40
        
        'Create Table
        .ListObjects.Add(xlSrcRange, Range("A1:H1"), , xlYes).Name = "Table" + YearName
    End With
    
    YearForm.Hide
End Sub
