VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddAccount 
   Caption         =   "Add an Account"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "AddAccount.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ButtonCancel_Click()
    AddAccount.Hide
End Sub

Private Sub ButtonSubmit_Click()
    MsgBox ("Placeholder Text")
End Sub

