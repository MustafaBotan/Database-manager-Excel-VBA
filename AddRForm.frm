VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddRForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6756
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8640.001
   OleObjectBlob   =   "AddRForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddRForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim i, nr As Integer
nr = WorksheetFunction.CountA(Columns("A:A")) - 1
Cells(1, 1).Select
ActiveCell.Offset(nr + 1, 0).Select
ActiveCell.Offset(0, 0) = AddRForm.TextBo1
ActiveCell.Offset(0, 1) = AddRForm.TextBo2
ActiveCell.Offset(0, 2) = AddRForm.TextBo3
ActiveCell.Offset(0, 3) = AddRForm.TextBo4
ActiveCell.Offset(0, 4) = AddRForm.TextBo5
ActiveCell.Offset(0, 5) = AddRForm.TextBo6
ActiveCell.Offset(0, 6) = AddRForm.TextBo7
ActiveCell.Offset(0, 7) = AddRForm.TextBo8
ActiveCell.Offset(0, 8) = AddRForm.TextBo9
ActiveCell.Offset(0, 9) = AddRForm.TextBo10
ActiveCell.Offset(0, 10) = AddRForm.TextBo11
ActiveCell.Offset(0, 11) = AddRForm.TextBo12

Unload AddRForm

End Sub

Private Sub CommandButton2_Click()
Unload AddRForm
End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label12_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub UserForm_Click()
Unload Me

End Sub
