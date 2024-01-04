VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim SelfIndex As Integer
Dim SelfCategory As String
Dim SelfTitle As String
Dim SelfLength As Integer
Dim SelfData(1) As Integer



Public Property Get Index() As Integer
    Index = SelfIndex
End Property

Public Property Let Index(ByVal NewIndex As Integer)
    SelfIndex = NewIndex
End Property


Public Property Get Category() As String
    Category = SelfCategory
End Property

Public Property Let Category(ByVal NewCategory As String)
    SelfCategory = NewCategory
End Property


Public Property Get Title() As String
    Title = SelfTitle
End Property

Public Property Let Title(ByVal NewTitle As String)
    SelfTitle = NewTitle
End Property


Public Property Get Length() As Integer
    Length = SelfLength
End Property

Public Property Let Length(ByVal NewLength As Integer)
    SelfLength = NewLength
End Property


