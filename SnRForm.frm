VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SnRForm 
   Caption         =   "UserForm1"
   ClientHeight    =   5304
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5940
   OleObjectBlob   =   "SnRForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SnRForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox2_Change()
SnRForm.Label4 = SnRForm.ComboBox2.Text
End Sub

Private Sub CommandButton1_Click()
Dim i, nr, colNmb As Integer
nr = Range("a1").CurrentRegion.Rows.Count
nc = Range("a1").CurrentRegion.Columns.Count

For i = 1 To nc
    If Range("B1:q1").Cells(1, i) = SnRForm.ComboBox2.Text Then
        colNmb = i
    End If
Next

For i = 1 To nr
    If Range("A1:A" & nr).Cells(i, 1) = SnRForm.ComboBox1.Text Then
        Range("A" & i).Select
        
        If Not ActiveCell.Offset(0, colNmb) = "" Then
            SnRForm.TextBox1 = ActiveCell.Offset(0, colNmb)
        Else
            If MsgBox(SnRForm.ComboBox2.Text & " is missing for " & SnRForm.ComboBox1.Text & vbNewLine & "Would you like to add this data?", vbYesNo) = vbYes Then
                ActiveCell.Offset(0, colNmb) = InputBox("Please enter " & ComboBox1.Text & "'s" & " new information for the " & SnRForm.ComboBox2.Text & " category.")
                TextBox1 = ActiveCell.Offset(0, colNmb)
            End If
            
        End If
        
    End If
Next i


End Sub

Private Sub CommandButton2_Click()
Dim i, nr, colNmb As Integer
nr = Range("a1").CurrentRegion.Rows.Count
nc = Range("a1").CurrentRegion.Columns.Count

For i = 1 To nc
    If Range("B1:q1").Cells(1, i) = SnRForm.ComboBox2.Text Then
        colNmb = i
    End If
Next

For i = 1 To nr
    If Range("A1:A" & nr).Cells(i, 1) = SnRForm.ComboBox1.Text Then
        Range("A" & i).Select
        
        ActiveCell.Offset(0, colNmb) = InputBox("Please enter new " & SnRForm.Label4 & " for " & SnRForm.ComboBox1.Text, , ActiveCell.Offset(0, colNmb))
        TextBox1 = ActiveCell.Offset(0, colNmb)
    End If
Next i
End Sub

Private Sub CommandButton3_Click()
Unload Me

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub
