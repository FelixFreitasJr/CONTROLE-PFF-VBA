VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "Login"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7890
   OleObjectBlob   =   "UserForm6.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cb_fechar_Click()

'FECHANDO A TELA
Unload Me

End Sub

Private Sub cb_logar_Click()

'CADASTRO PARA O LOGIN
    If txtlogin = "ADM" _
    And txtsenha = "1234" Then
'CASO O LOGIN POSITIVO
        Sheets("RELATORIO").Select
        Unload Me
        UserForm1.Hide
        Application.Visible = True
'CASO O LOGIN NEGATIVO
    Else
        MsgBox "Usuário ou Senha Incorretos!", vbInformation, "ERRO"
        txtlogin.Value = ""
        txtsenha.Value = ""
        UserForm6.txtlogin.SetFocus
    End If

End Sub

Private Sub txtlogin_Change()

'FORMATAÇÃO DOS CAMPOS EM CAIXA AUTA
    txtlogin = UCase(txtlogin.Value)

End Sub

Private Sub txtsenha_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

'FORMATAÇÃO PARA APENAS NUMEROS NA SENHA
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    
End Sub

Private Sub UserForm_Click()

End Sub
