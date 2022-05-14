VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Ultimate Organiser"
   ClientHeight    =   3048
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5976
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
AddCForm.Show
End Sub

Private Sub CommandButton2_Click()
Dim i, Nrows, Ncols As Integer

Range("A1").Select
Nrows = Range("a1").CurrentRegion.Rows.Count
For i = 1 To Nrows
    SnRForm.ComboBox1.AddItem Range("A1:A" & Nrows).Cells(i + 1, 1)
Next i
SnRForm.ComboBox1.Text = Range("a1:a" & Nrows).Cells(2, 1)

Range("A1").Select
Ncols = Range("a1").CurrentRegion.Columns.Count
For i = 1 To Ncols
    SnRForm.ComboBox2.AddItem Range("B1:Q1").Cells(1, i)
Next i
SnRForm.ComboBox2.Text = Range("B1:Q1").Cells(1, 1)

SnRForm.Show
End Sub

Private Sub CommandButton3_Click()
Dim i, Nrows As Integer

Range("A1").Select
Nrows = Range("a1").CurrentRegion.Rows.Count
For i = 1 To Nrows
    DeleteRForm.ComboBox1.AddItem Range("A1:A" & Nrows).Cells(i + 1, 1)
Next i
DeleteRForm.ComboBox1.Text = Range("a1:a" & Nrows).Cells(2, 1)
DeleteRForm.Show
End Sub

Private Sub CommandButton4_Click()
Dim i, Nrows As Integer

Range("A1").Select
Ncols = Sheet1.Cells(1, Columns.Count).End(xlToLeft).Column
For i = 1 To Ncols
    DeleteCForm.ComboBox1.AddItem Range("B1:Q1").Cells(1, i)
Next i
DeleteCForm.ComboBox1.Text = Range("B1:Q1").Cells(1, 1)
DeleteCForm.Show
End Sub

Private Sub CommandButton5_Click()
Dim nr, nc As Integer
Dim i As Integer
Dim categories() As Variant
nr = WorksheetFunction.CountA(Columns("A:A")) - 1
nc = Sheet1.Cells(1, Columns.Count).End(xlToLeft).Column
Cells(1, 1).Select
ReDim categories(nc) As Variant
For i = 1 To nc
    categories(i) = ActiveCell.Offset(0, i - 1)
Next i

With AddRForm
    .TextBo1.Visible = False
    .TextBo2.Visible = False
    .TextBo3.Visible = False
    .TextBo4.Visible = False
    .TextBo5.Visible = False
    .TextBo6.Visible = False
    .TextBo7.Visible = False
    .TextBo8.Visible = False
    .TextBo9.Visible = False
    .TextBo10.Visible = False
    .TextBo11.Visible = False
    .TextBo12.Visible = False
End With
If nc >= 1 Then
    AddRForm.Labe1 = categories(1)
    AddRForm.TextBo1.Visible = True
        AddRForm.Width = 250
End If
If nc >= 2 Then
    AddRForm.Labe2 = categories(2)
    AddRForm.TextBo2.Visible = True
        AddRForm.Width = 250
End If
If nc >= 3 Then
    AddRForm.Labe3 = categories(3)
    AddRForm.TextBo3.Visible = True
        AddRForm.Width = 250
End If
If nc >= 4 Then
    AddRForm.Labe4 = categories(4)
    AddRForm.TextBo4.Visible = True
        AddRForm.Width = 250
End If
If nc >= 5 Then
    AddRForm.Labe5 = categories(5)
    AddRForm.TextBo5.Visible = True
    AddRForm.Width = 250
End If
If nc >= 6 Then
    AddRForm.Labe6 = categories(6)
    AddRForm.TextBo6.Visible = True
    AddRForm.Width = 250
End If
If nc >= 7 Then
    AddRForm.Labe7 = categories(7)
    AddRForm.TextBo7.Visible = True
        AddRForm.Width = 450
End If
If nc >= 8 Then
    AddRForm.Labe8 = categories(8)
    AddRForm.TextBo8.Visible = True
    AddRForm.Width = 450
End If
If nc >= 9 Then
    AddRForm.Labe9 = categories(9)
    AddRForm.TextBo9.Visible = True
    AddRForm.Width = 450
End If
If nc >= 10 Then
    AddRForm.Labe10 = categories(10)
    AddRForm.TextBo10.Visible = True
    AddRForm.Width = 450
End If
If nc >= 11 Then
    AddRForm.Labe11 = categories(11)
    AddRForm.TextBo11.Visible = True
    AddRForm.Width = 450
End If
If nc >= 12 Then
    AddRForm.Labe12 = categories(12)
    AddRForm.TextBo12.Visible = True
    AddRForm.Width = 450
End If

AddRForm.Show

End Sub

Private Sub CommandButton6_Click()
Unload Me
End Sub
