VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CostEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim SelfID As Integer
Dim SelfYear As Integer
Dim SelfMonth As Integer
Dim SelfDay As Integer
Dim SelfDate As Date
Dim SelfCost As Currency
Dim SelfPlace As String
Dim SelfLocation As String
Dim SelfMethod As String
Dim SelfNotes As String


Public Property Get EntryDate() As Date
    EntryDate = SelfDate
End Property


Private Property Let EntryDate(ByVal NewDate As Date)
    SelfDate = NewDate
End Property


Public Property Get EntryCost() As Currency
    EntryCost = SelfCost
End Property


Public Property Let EntryCost(ByVal NewCost As Currency)
    SelfCost = NewCost
End Property


Public Property Get EntryPlace() As String
    EntryPlace = SelfPlace
End Property


Public Property Let EntryPlace(ByVal NewPlace As String)
    SelfPlace = NewPlace
End Property


Public Property Get EntryLocation() As String
    EntryLocation = SelfLocation
End Property


Public Property Let EntryLocation(ByVal NewLocation As String)
    SelfLocation = NewLocation
End Property


Public Property Get EntryMethod() As String
    EntryMethod = SelfMethod
End Property


Public Property Let EntryMethod(ByVal NewMethod As String)
    SelfMethod = NewMethod
End Property


Public Property Get EntryNotes() As String
    EntryNotes = SelfNotes
End Property


Public Property Let EntryNotes(ByVal NewNotes As String)
    SelfNotes = NewNotes
End Property
