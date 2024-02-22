VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CostForm 
   Caption         =   "Cost Entry Form"
   ClientHeight    =   7080
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


Private Sub UserForm_Initialize()
    Call PopulateID
    Call PopulateDateBoxes
    Call PopulateMethodBox
    Call PopulateCategoryBox
End Sub


Private Sub PopulateID()
    Dim MySheet As Worksheet
    Dim ID As Integer
    
    Set MySheet = Application.Worksheets("Inputs")
    Let ID = MySheet.Range("J4").Value
    ID = ID + 1
    
    CostForm.IDBox.Value = ID
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


Private Sub PopulateMethodBox()
    Dim MySheet As Worksheet
    Dim MyTable As ListObject
    Dim MyRow As ListRow
    Dim MyRange As Range
    Dim MyString As String
    
    Set MySheet = Application.Worksheets("Inputs")
    Set MyTable = MySheet.ListObjects("MethodTable")
    
    For Each MyRow In MyTable.ListRows:
        Set MyRange = MyRow.Range(1, 1)
        Let MyString = MyRange.Value
        CostForm.MethodBox.AddItem (MyString)
    Next
End Sub


Private Sub MethodBox_AfterUpdate()
    'check here to ensure option is in MethodsTable
    Dim InputMethod As String
    Dim Result As Boolean
    Dim MethodRow As ListRow
    Dim MethodTable As ListObject
    
    Result = False
    InputMethod = CostForm.MethodBox.Value
    Set MethodTable = ActiveSheet.ListObjects("MethodTable")
    
    For Each MethodRow In MethodTable.ListRows:
        If MethodRow.Range(1, 1).Value = InputMethod Then
            Result = True
        End If
    Next
    
    If Result = False Then
        MsgBox ("Please select a valid input!")
        CostForm.MethodBox.ForeColor = RGB(255, 0, 0)
    End If
    
    If Result = True Then
        CostForm.MethodBox.ForeColor = RGB(0, 0, 0)
    End If
    
End Sub


Private Sub PopulateCategoryBox()
    Dim MySheet As Worksheet
    Dim MyTable As ListObject
    Dim MyRow As ListRow
    Dim MyRange As Range
    Dim MyString As String
    
    Set MySheet = Application.Worksheets("Inputs")
    
    Set MyTable = MySheet.ListObjects("CategoryTable")
    
    For Each MyRow In MyTable.ListRows:
        Set MyRange = MyRow.Range(1, 1)
        Let MyString = MyRange.Value
        CostForm.CategoryBox.AddItem (MyString)
    Next
End Sub


Private Sub CategoryBox_AfterUpdate()
    'Check here to ensure option is in CategoryTable
    Dim InputCategory As String
    Dim Result As Boolean
    Dim CategoryRow As ListRow
    Dim CategoryTable As ListObject
    
    Result = False
    InputCategory = CostForm.CategoryBox.Value
    Set CategoryTable = ActiveSheet.ListObjects("CategoryTable")
    
    For Each CategoryRow In CategoryTable.ListRows:
        If CategoryRow.Range(1, 1).Value = InputCategory Then
            Result = True
        End If
    Next
    
    If Result = False Then
        MsgBox ("Please select a valid input!")
        CostForm.CategoryBox.ForeColor = RGB(255, 0, 0)
    End If
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
                    CategoryBox = LRow.Range(1, 6)
                    MethodBox = LRow.Range(1, 7)
                    NotesBox = LRow.Range(1, 8)
                End If
                
            Next LRow
        End If
    Next WS
End Sub


Private Sub AddButton_Click()
    ' Find sheet for input year
    ' Create New Row
    ' Plug info into New Row
    
    Dim WSName As String
    Dim WS As Worksheet
    
    Dim LRow As ListRow
    Dim IDValue As String
    
    Dim IDYear As String
    Dim IDMonth As String
    Dim IDDay As String
    
    Dim IDDateString As String
    Dim IDDate As Date
    
    WSName = CStr(CostForm.YearBox.Value)
    Set WS = Application.Worksheets(WSName)
    
    IDValue = CStr(CostForm.IDBox.Value)
    IDYear = CStr(CostForm.YearBox.Value)
    IDMonth = CStr(CostForm.MonthBox.Value)
    IDDay = CStr(CostForm.DayBox.Value)
    IDDateString = IDMonth + "-" + IDDay + "-" + IDYear
    IDDate = CDate(IDDateString)
    
    
    Set LRow = WS.ListObjects("Table" + WS.Name).ListRows.Add
    
    With LRow
        .Range(1, 1).Value = IDValue
        .Range(1, 2).Value = IDDate
        .Range(1, 3).Value = CostForm.CostBox.Value
        .Range(1, 4).Value = CostForm.PlaceBox.Value
        .Range(1, 5).Value = CostForm.LocationBox.Value
        .Range(1, 6).Value = CostForm.CategoryBox.Value
        .Range(1, 7).Value = CostForm.MethodBox.Value
        If CostForm.NotesBox.Value <> "(Optional)" Then
            .Range(1, 8).Value = CostForm.NotesBox.Value
        End If
    End With
    
    Application.Worksheets("Inputs").Range("J4:J4").Value = CInt(IDValue)
    MsgBox ("Entry Added")
    Call ResetButton_Click
    
End Sub


Private Sub EditButton_Click()
    Dim WSName As String
    Dim WS As Worksheet
    
    Dim LRow As ListRow
    Dim IDValue As String
    
    Dim IDYear As String
    Dim IDMonth As String
    Dim IDDay As String
    
    Dim IDDateString As String
    Dim IDDate As Date
    
    WSName = CStr(CostForm.YearBox.Value)
    Set WS = Application.Worksheets(WSName)
    
    IDValue = CStr(CostForm.IDBox.Value)
    IDYear = CStr(CostForm.YearBox.Value)
    IDMonth = CStr(CostForm.MonthBox.Value)
    IDDay = CStr(CostForm.DayBox.Value)
    IDDateString = IDMonth + "-" + IDDay + "-" + IDYear
    IDDate = CDate(IDDateString)
    
    For Each LRow In WS.ListObjects("Table" + WS.Name).ListRows
                           
        If LRow.Range(1, 1).Value = IDValue Then

            LRow.Range(1, 1) = IDValue
            LRow.Range(1, 2) = IDDate
            LRow.Range(1, 3) = CostForm.CostBox.Value
            LRow.Range(1, 4) = CostForm.PlaceBox.Value
            LRow.Range(1, 5) = CostForm.LocationBox.Value
            LRow.Range(1, 6) = CostForm.CategoryBox.Value
            LRow.Range(1, 7) = CostForm.MethodBox.Value
            LRow.Range(1, 8) = CostForm.NotesBox.Value
        End If
                
    Next LRow
    MsgBox ("Entry Edited")

End Sub


Private Sub DeleteButton_Click()
   
    Dim WSName As String
    Dim WS As Worksheet
    Dim LRow As ListRow
    Dim IDValue As String

    WSName = CStr(CostForm.YearBox.Value)
    Set WS = Application.Worksheets(WSName)
    IDValue = CStr(CostForm.IDBox.Value)
    
    
    For Each LRow In WS.ListObjects("Table" + WS.Name).ListRows
                           
        If CStr(LRow.Range(1, 1).Value) = IDValue Then

            Call LRow.Delete
        
        End If
                
    Next LRow
   
    MsgBox ("Entry Deleted")
    Call ResetButton_Click
   
End Sub


Private Sub ResetButton_Click()
    Call ClearIDSearch
    Call PopulateDateBoxes
    Call PopulateID
    
    CostForm.CostBox.Value = ""
    CostForm.PlaceBox.Value = ""
    CostForm.LocationBox.Value = ""
    CostForm.CategoryBox.Value = ""
    CostForm.MethodBox.Value = ""
    CostForm.NotesBox.Value = ""
    
End Sub




