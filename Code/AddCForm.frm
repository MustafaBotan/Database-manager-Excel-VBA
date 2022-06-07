VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddCForm 
   Caption         =   "UserForm2"
   ClientHeight    =   2472
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5184
   OleObjectBlob   =   "AddCForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddCForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim i As Integer

Sheets("sheet1").Select

For i = 1 To 99
    If Cells(1, i) = "" Then
        Cells(1, i) = AddCForm.TextBox1
        AddCForm.Hide
        Exit Sub
    End If
Next
AddCForm.Hide
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub
