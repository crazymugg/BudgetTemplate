VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Accounts table
'
' Account
' Name - str
' Type - int
'
'


Option Explicit

Dim SelfName As String
Dim SelfType As String


Public Property Get Name() As String
    Name = SelfName
End Property

Public Property Let Name(ByVal NewName As String)
    SelfName = NewName
End Property


Public Property Get AccType() As String
    AccType = SelfType
End Property

Public Property Let AccType(ByVal NewType As String)
    SelfType = NewType
End Property
