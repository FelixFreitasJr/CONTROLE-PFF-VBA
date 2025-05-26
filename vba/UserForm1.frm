VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Controle PFF 3.7 - Externo"
   ClientHeight    =   8460.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19140
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Fechar As Boolean

Private Sub btbusca_Click() 'LIMPANDO CAMPOS

'LIMPANDO TODOS OS CAMPOS
    Sheets("dados").Range("G2").Value = ""
    txtcpf.Text = ""
    txtnome.Text = ""
    txtempresa.Text = ""
    txtfuncao.Text = ""
    cboxjustificativa.Text = " "
    txtstatus.Text = ""
    txtvalid.Text = ""
    Frame1.BackColor = &HFFFFFF
    TextBox1.Text = ""
    TextBox2.Text = ""
    TextBox3.Text = ""
    txtnome.BackColor = &HFFFFFF
    txtnome.ForeColor = &H8000000D
    TextBox1.ForeColor = &H8000000D
    txtempresa.BackColor = &HFFFFFF
    txtempresa.ForeColor = &H8000000D
    TextBox2.ForeColor = &H8000000D
    txtfuncao.BackColor = &HFFFFFF
    txtfuncao.ForeColor = &H8000000D
    TextBox3.ForeColor = &H8000000D
    TextBox3.BackColor = &HFFFFFF
    TextBox2.BackColor = &HFFFFFF
    TextBox1.BackColor = &HFFFFFF

'FOCANDO CURSOR
    txtcpf.SetFocus
End Sub

Private Sub cb_fechar_form_Click() 'CHAMANDO ACESSO AO CONTROLE DO EXCEL

'CHAMANDO ACESSSO AO ADMINISTRATIVO
    UserForm6.Show

End Sub

Private Sub cb_pesq_Click()
txtcpf.Text = ""
UserForm5.Show


End Sub

Private Sub cbdispensar_Click() 'INFORMANDO OS DADOS NA PLANILHA

'VALIDAÇÃO DE CAMPO VAZIO
    If cboxjustificativa.Value = "" Or cboxjustificativa.Value = " " Or txtnome.Value = "CPF NÃO CADASTRADO" Then
        MsgBox "Não Localizado, favor verificar.", vbCritical, "Erro!"
        cboxjustificativa.SetFocus
    
'VALIDAÇÃO DE FALTA INFORMAÇÃO
    ElseIf txtempresa.Value = "NÃO INFORMADO" Or txtfuncao.Value = "NÃO INFORMADO" Then
        MsgBox "Favor atualizar o cadastro", vbCritical, "Erro!"
        UserForm4.Show
    Else

'DISPENSANDO CASO ESTEJA TUDO OK
    Call dispensar

'SALVANDO OS DADOS
    Call salvar
    
'LIMPANDO ANTES DE SALVAR
    txtcpf.Text = ""
    txtnome.Text = ""
    txtempresa.Text = ""
    txtfuncao.Text = " "
    cboxjustificativa.Text = " "
    txtstatus.Text = ""
    txtobs.Text = ""
    txtvalid.Text = ""
    TextBox1.Text = ""
    TextBox2.Text = ""
    TextBox3.Text = ""
    Frame1.BackColor = &HFFFFFF
    txtnome.BackColor = &HFFFFFF
    txtnome.ForeColor = &H8000000D
    TextBox1.ForeColor = &H8000000D
    txtempresa.BackColor = &HFFFFFF
    txtempresa.ForeColor = &H8000000D
    TextBox2.ForeColor = &H8000000D
    txtfuncao.BackColor = &HFFFFFF
    txtfuncao.ForeColor = &H8000000D
    TextBox3.ForeColor = &H8000000D
    TextBox3.BackColor = &HFFFFFF
    TextBox2.BackColor = &HFFFFFF
    TextBox1.BackColor = &HFFFFFF


'FOCANDO O CURSOR
    txtcpf.SetFocus
   
End If

End Sub

Private Sub cbfechar_Click()

'SALVANDO ARQUIVO
    Call salvar
'FECHANDO O ARQUIVO
    Application.Quit

End Sub

Private Sub cbinformacao_Click()

'ABRIR DADOS
    UserForm2.Show

End Sub

Private Sub cbnovocadastro_Click()

'VALIDANDO CPF PARA CADASTRO
    If txtcpf.Value = "" Or txtcpf.Value = " " Then
        MsgBox "Favor Informar o CPF", vbCritical, "Erro!"
        txtcpf.SetFocus

'CPF VALIDADO ABRINDO CADASTRO
    ElseIf Sheets("dados").Range("N2").Value = True Then
        UserForm4.Show
'CPF INVALIDO INFORMATIVO
    Else
        MsgBox "CPF Inválido, Digite novamente!", vbCritical, "ERRO!"
End If

End Sub

Private Sub cboxjustificativa_Change()

'CONTROLE DE SAIDA PARA VISITANTE
    If UserForm1.cboxjustificativa.Value = "VISITANTE" Then
        UserForm1.txtquant.Enabled = True
        UserForm1.Label15.Enabled = True
        UserForm1.txtquant.BackColor = &HFF& 'VERMELHO
        UserForm1.txtquant.ForeColor = &HFFFFFF  'BRANCO
        UserForm1.txtquant.Text = "1"
    Else
        UserForm1.txtquant.Enabled = False
        UserForm1.Label15.Enabled = False
        UserForm1.txtquant.BackColor = &HFFFFFF 'BRANCO
        UserForm1.txtquant.ForeColor = &HFF& 'VERMELHO
        UserForm1.txtquant.Text = "1"
End If

End Sub

Private Sub cbrelatorio_Click()

'RELATORIO - EM DESESENVOLVIMENTO
   UserForm3.Show

End Sub

Private Sub CommandButton1_Click()

End Sub




Private Sub txtcpf_Change()

'FORMATAÇÃO DO CAMPO PARA OBTER OS PONTOS E TRAÇOS DO CPF
    Dim CPF As String, CPF2 As String, CPF3 As String
    Dim i As Integer, j As Integer, n As Integer

CPF = txtcpf.Value
i = Len(CPF)

    For j = 1 To i
        If IsNumeric(Mid(CPF, j, 1)) Then
            CPF2 = CPF2 & Mid(CPF, j, 1)
        End If
   Next

i = Len(CPF2)
    For j = 1 To i
        CPF3 = CPF3 & Mid(CPF2, j, 1)
       If j = 4 Or j = 7 Then
            n = Len(CPF3) - 1
            CPF3 = Left(CPF3, n) & "." & Right(CPF3, 1)
            ElseIf j = 10 Then
            n = Len(CPF3) - 1
            CPF3 = Left(CPF3, n) & "-" & Right(CPF3, 1)

           End If

       Next
       
    txtcpf.Value = CPF3

'FORMATAÇÃO VISUAL DOS CAMPOS
    Frame1.BackColor = &HFFFFFF ' BRANCO
    Frame2.BackColor = &HFFFFFF ' BRANCO

'FORMATAÇÃO VISUAL DOS CAMPOS CASO "NÃO INFORMADO"
  
    
End Sub

Private Sub txtcpf_Exit(ByVal Cancel As MSForms.ReturnBoolean)

'BUSCAR PROCV NA PLANILHA
    Sheets("dados").Range("G2").Value = txtcpf.Value
    Sheets("dados").Range("G6").Value = txtcpf.Value
    Sheets("dados").Range("H6").Value = txtnome.Value

'MENSAGEM SE ERRO DE CPF
If Sheets("dados").Range("N2").Value = False Then
        MsgBox "CPF Inválido, Verifique e Digite novamente!", vbCritical, "ERRO!"
        txtcpf.SetFocus
    End If

'CASO ERRO PROSSIGA
    On Error Resume Next
        txtnome.Value = Sheets("dados").Range("H2").Value
    On Error Resume Next
        txtempresa.Value = Sheets("dados").Range("I2").Value
    On Error Resume Next
        txtfuncao.Value = Sheets("dados").Range("J2").Value
    On Error Resume Next
        txtstatus.Value = Sheets("dados").Range("K2").Value
    On Error Resume Next
        txtvalid.Value = Sheets("dados").Range("L2").Value
    On Error Resume Next
        TextBox1.Value = Sheets("dados").Range("H1").Value
    On Error Resume Next
        TextBox2.Value = Sheets("dados").Range("I1").Value
    On Error Resume Next
        TextBox3.Value = Sheets("dados").Range("J1").Value
    On Error Resume Next
        txtcontr.Value = Sheets("dados").Range("M2").Value
    On Error Resume Next
        Textsoli.Value = Sheets("dados").Range("M1").Value
    On Error Resume Next
        Textstatus.Value = Sheets("dados").Range("K1").Value
    On Error Resume Next
        Textvenc.Value = Sheets("dados").Range("L1").Value
    On Error Resume Next
        Text_loc01.Value = Sheets("dados").Range("M4").Value
    On Error Resume Next
        Text_loc02.Value = Sheets("dados").Range("M5").Value

'STATUS DA DISPENSA
    If txtstatus.Text = " " Then
        Frame1.BackColor = &HFFFFFF
    End If

    If txtstatus.Text = "LIBERADO" Then
        Frame1.BackColor = &HC000&
        txtstatus.ForeColor = &HFFFFFF
    End If

    If txtstatus.Text = "ANTES DO PRAZO" Then
        Frame1.BackColor = &HFF&
        txtstatus.ForeColor = &HFFFFFF
    End If

'DADOS PESSOAIS NÃO CADASTRADOS OU DESATUALIZADOS
    If txtnome.Text = "CPF NÃO CADASTRADO" Then
        Frame2.BackColor = &HFF& 'VERMELHO
        txtnome.ForeColor = &HFFFFFF 'BRANCO
        TextBox1.ForeColor = &HFFFFFF 'BRANCO
        UserForm1.TextBox1.BackColor = &HFF&    'VERMELHO
        UserForm1.txtnome.BackColor = &HFF&     'VERMELHO
        UserForm1.cbnovocadastro.SetFocus
    Else
        Frame2.BackColor = &HFFFFFF 'BRANCO
        txtnome.ForeColor = &H8000000D 'AZUL
        TextBox1.ForeColor = &H8000000D 'AZUL
        UserForm1.TextBox1.BackColor = &HFFFFFF    'BRANCO
        UserForm1.txtnome.BackColor = &HFFFFFF     'VERMELHO
    End If

    If txtempresa.Text = "CPF NÃO CADASTRADO" Then
        Frame2.BackColor = &HFF& 'VERMELHO
        txtempresa.ForeColor = &HFFFFFF 'BRANCO
        TextBox2.ForeColor = &HFFFFFF 'BRANCO
        UserForm1.TextBox2.BackColor = &HFF&    'VERMELHO
        UserForm1.txtempresa.BackColor = &HFF&  'VERMELHO
        UserForm1.cbnovocadastro.SetFocus
    Else
        Frame2.BackColor = &HFFFFFF 'BRANCO
        txtempresa.ForeColor = &H8000000D 'AZUL
        TextBox2.ForeColor = &H8000000D 'AZUL
        UserForm1.TextBox2.BackColor = &HFFFFFF    'BRANCO
        UserForm1.txtempresa.BackColor = &HFFFFFF  'BRANCO
    End If

    If txtfuncao.Text = "CPF NÃO CADASTRADO" Then
        Frame2.BackColor = &HFF& 'VERMELHO
        txtfuncao.ForeColor = &HFFFFFF 'BRANCO
        TextBox3.ForeColor = &HFFFFFF 'BRANCO
        UserForm1.TextBox3.BackColor = &HFF&    'VERMELHO
        UserForm1.txtfuncao.BackColor = &HFF&  'VERMELHO
        UserForm1.cbnovocadastro.SetFocus
    Else
        Frame2.BackColor = &HFFFFFF 'BRANCO
        txtfuncao.ForeColor = &H8000000D 'AZUL
        TextBox3.ForeColor = &H8000000D 'AZUL
        UserForm1.TextBox3.BackColor = &HFFFFFF    'BRANCO
        UserForm1.txtfuncao.BackColor = &HFFFFFF  'BRANCO
    End If
    
    If UserForm1.txtempresa.Text = "NÃO INFORMADO" Then
        UserForm1.txtempresa.BackColor = &HFF& 'VERMELHO
        UserForm1.txtempresa.ForeColor = &HFFFFFF     'BRANCO
        UserForm1.TextBox2.BackColor = &HFF&    'VERMELHO
        UserForm1.TextBox2.ForeColor = &HFFFFFF     'BRANCO
        UserForm1.cbnovocadastro.SetFocus
    End If
    
    If UserForm1.txtfuncao.Text = "NÃO INFORMADO" Then
        UserForm1.txtfuncao.BackColor = &HFF& 'VERMELHO
        UserForm1.txtfuncao.ForeColor = &HFFFFFF      'BRANCO
        UserForm1.TextBox3.BackColor = &HFF&    'VERMELHO
        UserForm1.TextBox3.ForeColor = &HFFFFFF     'BRANCO
        UserForm1.cbnovocadastro.SetFocus
    End If
    
End Sub

Private Sub txtcpf_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

'CONTROLE DE ENTRADA DE APENAS NUMEROS NO CAMPO CPF
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub




 Sub UserForm_Initialize()

'FOCANDO O CURSO NO INICIO DA OPERAÇÃO
    txtcpf.SetFocus
    
'FORMATAÇÃO DOS CAMPOS EM CAIXA AUTA
    lbdata = UCase(FORMAT(Date, "dd mmmm yyyy"))
    lbhora = Time
    txtnome = UCase(txtnome.Value)
    txtempresa = UCase(txtempresa.Value)
    txtfuncao = UCase(txtfuncao.Value)
    UserForm1.Label14 = UCase(Application.UserName)

'FORMATAÇÃO DAS JUSTIFICATIVAS DA DISPENSA
    Dim just As Integer
        just = Range("DADOS!E1").End(xlDown).Row
        cboxjustificativa.RowSource = "DADOS!E1:E" & just

'INICIO DA CONTAGEM DO RELOGIO DO MONITOR
    Call IniciarRelogio

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

'BLOQUEIO DE FECHAMENTO ERRADO DA TELA
    If Fechar = False Then
        MsgBox "Clique no botão fechar do formulário!", vbCritical, "FECHAR"
        Cancel = True
    End If

End Sub
