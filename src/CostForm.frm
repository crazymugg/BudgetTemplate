VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CostForm 
   Caption         =   "Cost Entry Form"
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
    'check here to ensure option is in MethodsTable
    Dim InputMethod As String
    Dim Result As Boolean
    Dim MethodRow As ListRow
    Dim MethodTable As ListObject
    
    Result = False
    InputMethod = CostForm.MethodBox.Value
    Set MethodTable = ActiveSheet.ListObjects("MethodsTable")
    
    For Each MethodRow In MethodTable.ListRows:
        If MethodRow.Range(1, 1).Value = InputMethod Then
            Result = True
        End If
    Next
    
    If Result = False Then
        MsgBox ("Please select a valid input!")
        CostForm.MethodBox.ForeColor = RGB(255, 0, 0)
    End If
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
        CostForm.MethodBox.AddItem (MyString)
    Next
End Sub


Private Sub ClearIDSearch()
    CostForm.IDSearchBox.Value = ""
End Sub


Private Sub SearchButton_Click()
    Dim WS As Worksheet
    Dim ORow As Range
    Dim LRow As ListRow
    Dim IDValue As String
    Dim IDDate As Date
    
    For Each WS In Application.Worksheets
    
        If WS.Name = "Inputs" Or WS.Name = "Accounts" Then
            Debug.Print ("Skip this sheet: " + WS.Name)
        Else
            
            For Each LRow In WS.ListObjects("Table" + WS.Name).ListRows
            
                IDValue = CStr(LRow.Range(1, 1).Value)
                
                If IDValue = IDSearchBox.Value Then
                    CostForm.IDBox = IDValue
                    IDDate = LRow.Range(1, 2)
                    YearBox = Year(IDDate)
                    MonthBox = Month(IDDate)
                    DayBox = Day(IDDate)
                    CostBox = LRow.Range(1, 3)
                    PlaceBox = LRow.Range(1, 4)
                    LocationBox = LRow.Range(1, 5)
                    MethodBox = LRow.Range(1, 6)
                    NotesBox = LRow.Range(1, 7)
                End If
                
            Next LRow
        End If
    Next WS

End Sub


Private Sub AddButton_Click()
    ' Find sheet for input year
    ' Create New Row
    ' Plug info into New Row
End Sub


Private Sub EditButton_Click()
    ' Find sheet for input year
    ' Find Row for ID
    ' Override info in row
    
    Dim WS As Worksheet
    Dim ORow As Range
    Dim LRow As ListRow
    Dim IDValue As String
    Dim IDDate As Date
    
    For Each WS In Application.Worksheets
    
        If WS.Name = "Inputs" Or WS.Name = "Accounts" Then
            Debug.Print ("Skip this sheet: " + WS.Name)
        Else
            
            For Each LRow In WS.ListObjects("Table" + WS.Name).ListRows
            
                IDValue = CStr(LRow.Range(1, 1).Value)
                
                If IDValue = IDSearchBox.Value Then
                    
                
                
                
                    CostForm.IDBox = IDValue
                    IDDate = LRow.Range(1, 2)
                    YearBox = Year(IDDate)
                    MonthBox = Month(IDDate)
                    DayBox = Day(IDDate)
                    CostBox = LRow.Range(1, 3)
                    PlaceBox = LRow.Range(1, 4)
                    LocationBox = LRow.Range(1, 5)
                    MethodBox = LRow.Range(1, 6)
                    NotesBox = LRow.Range(1, 7)
                End If
                
            Next LRow
        End If
    Next WS
End Sub


Private Sub DeleteButton_Click()
   ' Find Sheet for input year
   ' Find Row for ID
   ' Delete Row
End Sub


Private Sub ResetButton_Click()
    Call ClearIDSearch
    Call PopulateDateBoxes
    Call PopulateID
    
    CostForm.CostBox.Value = ""
    CostForm.PlaceBox.Value = ""
    CostForm.LocationBox.Value = ""
    CostForm.MethodBox.Value = ""
    CostForm.NotesBox.Value = ""
    
End Sub




