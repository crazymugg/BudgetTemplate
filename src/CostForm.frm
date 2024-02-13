VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CostForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   OleObjectBlob   =   "CostForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CostForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub MethodBox_AfterUpdate()
     MsgBox ("Hi")
     'TODO add a check here to ensure option is in MethodsTable
End Sub


Private Sub UserForm_Initialize()
    Call PopulateID
    Call PopulateDateBoxes
    Call PopulateMethodBox
End Sub


Private Sub PopulateDateBoxes()
    Dim MyDate As Date
    Dim MyDateString As String
    Dim DateArray() As String
    
    MyDate = Date
    MyDateString = CStr(MyDate)
    
    DateArray = Split(MyDateString, "-")
    
    CostForm.YearBox.Value = DateArray(0)
    CostForm.MonthBox.Value = DateArray(1)
    CostForm.DayBox.Value = DateArray(2)
End Sub


Private Sub PopulateID()
    Dim MySheet As Worksheet
    Dim ID As Integer
    
    Set MySheet = Application.Worksheets("Inputs")
    Let ID = MySheet.Range("J4").Value
    ID = ID + 1
    
    CostForm.IDBox.Value = ID
    
End Sub


Private Sub PopulateMethodBox()
    Dim MySheet As Worksheet
    Dim MyTable As ListObject
    Dim MyRow As ListRow
    Dim MyRange As Range
    Dim MyString As String
    
    Set MySheet = Application.Worksheets("Inputs")
    
    Set MyTable = MySheet.ListObjects("MethodsTable")
    
    For Each MyRow In MyTable.ListRows:
        Set MyRange = MyRow.Range(1, 1)
        Let MyString = MyRange.Value
        Debug.Print ("String is" + MyString)
        CostForm.MethodBox.AddItem (MyString)
    Next
End Sub


Private Sub ClearIDSearch()
    CostForm.IDSearchBox.Value = ""
End Sub


Private Sub AddButton_Click()
    Debug.Print ("Add Button Pressed")
End Sub


Private Sub EditButton_Click()
    Debug.Print ("Edit Button Pressed")
End Sub


Private Sub DeleteButton_Click()
    Debug.Print ("Delete Button Pressed")
End Sub


Private Sub ResetButton_Click()
    Call ClearIDSearch
    Call PopulateDateBoxes
    Call PopulateID
    
    CostForm.IDBox.Value = ""
    CostForm.CostBox.Value = ""
    CostForm.PlaceBox.Value = ""
    CostForm.LocationBox.Value = ""
    CostForm.MethodBox.Value = ""
    CostForm.NotesBox.Value = ""
    
End Sub




