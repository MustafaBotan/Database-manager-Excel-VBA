VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteRForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "DeleteRForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeleteRForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

End Sub

Private Sub ComboBox2_Change()
Dim i, Nrows As Integer, j As Integer
DeleteRForm.ComboBox1.Clear
Range("A1").Select
Nrows = Range("a1").CurrentRegion.Rows.Count

For i = 1 To Nrows
    DeleteRForm.ComboBox1.AddItem ActiveCell.Offset(i, 0)
Next i
End Sub

Private Sub CommandButton1_Click()
Dim i, nr As Integer
nr = Range("a1").CurrentRegion.Rows.Count
For i = 1 To nr
    If Range("A1:A" & nr).Cells(i, 1) = DeleteRForm.ComboBox1.Text Then
        Range("A" & i).Select
        ans = MsgBox("Are you sure you want to delete the row", vbYesNo)
        If ans = 6 Then
        Rows(i).Delete
        End If
    End If
Next i

If ans = 6 Then
Unload DeleteRForm
End If
End Sub

Private Sub CommandButton2_Click()
Unload Me

End Sub
