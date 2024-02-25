VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "..."
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17235
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub btlimpar157_Click()
UserForm5.TextBox1.Value = ""
UserForm5.TextBox1.SetFocus
End Sub

Private Sub cbfechar23_Click()

    Unload Me

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim nLin As Integer

nLin = ListBox1.ListIndex

If nLin = -1 Then Exit Sub
If ListBox1.Value = 0 Then Exit Sub

Sheets("dados").Range("G5").Value = UserForm5.ListBox1.List(nLin, 0)
On Error Resume Next
        UserForm1.txtcpf.Value = Sheets("dados").Range("G5").Value
        Sheets("dados").Range("G5").Value = ""
Unload Me
UserForm1.txtcpf.SetFocus
End Sub


Private Sub TextBox1_AfterUpdate()
Dim A As String, Coluna As Integer, LinhaListBox As Integer, Linha As Integer

Linha = 2
LinhaListBox = 0
With Sheets("CADASTRO").Select
With Me.ListBox1
.Clear
While Cells(Linha, 1) <> Empty
For Coluna = 1 To 4
A = Cells(Linha, Coluna)
If InStr(1, UCase(A), UCase(Me.TextBox1.Text)) > 0 Then
.AddItem
.List(LinhaListBox, 0) = Cells(Linha, 1)
.List(LinhaListBox, 1) = Cells(Linha, 2)
.List(LinhaListBox, 2) = Cells(Linha, 3)
.List(LinhaListBox, 3) = Cells(Linha, 4)
LinhaListBox = LinhaListBox + 1

GoTo proxima_linha
End If

Next Coluna
proxima_linha:
Linha = Linha + 1
Wend
Me.Label1.Object = .ListCount & " Registro(s) encontrados(s)"
End With
End With
End Sub



Private Sub TextBox1_Change()
UserForm5.TextBox1.Text = UCase(TextBox1.Value)
End Sub

Private Sub UserForm_Activate()

End Sub

Private Sub UserForm_Click()

End Sub
