VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteCForm 
   Caption         =   "UserForm1"
   ClientHeight    =   2640
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4788
   OleObjectBlob   =   "DeleteCForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeleteCForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()
Dim i, nc As Integer
Dim ans As Integer
nc = Sheet1.Cells(1, Columns.Count).End(xlToLeft).Column
For i = 1 To nc
    If Range("B1:q1").Cells(1, i) = DeleteCForm.ComboBox1.Text Then
        ans = MsgBox("Are you sure you want to delete the category", vbYesNo)
        If ans = 6 Then
        Columns(i + 1).Delete
        End If
    End If
Next i

If ans = 6 Then
Unload DeleteCForm
End If

End Sub

Private Sub CommandButton2_Click()
Unload Me

End Sub

