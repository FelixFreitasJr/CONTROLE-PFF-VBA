VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "Controle PFF 3.7"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11625
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBCADASTRO_Change()

'FORMATAÇÃO PARA OS CAMPOS DO FORMULARIO DE CADASTRO

'SE O CAMPO ESTVER COMO "VISITANTE"
    
'SE POSITIVO, BLOQUEIO DE CAMPOS E PREENCHIMENTO AUTOMARICO
    If UserForm4.CBCADASTRO.Value = "VISITANTE" Then
        
        UserForm4.txtnome2.Enabled = True
        UserForm4.txtnome2.Text = ""
        
        UserForm4.txtempresa2.Enabled = False
        UserForm4.txtempresa2.Text = "VISITANTE"
        UserForm4.txtempresa2.BackColor = &HFFFFFF 'BRANCO
        UserForm4.txtempresa2.ForeColor = &H8000000D 'AZUL
        
        UserForm4.txtfuncao2.Enabled = False
        UserForm4.txtfuncao2.Text = "VISITANTE"
        UserForm4.txtfuncao2.BackColor = &HFFFFFF 'BRANCO
        UserForm4.txtfuncao2.ForeColor = &H8000000D 'AZUL
        
    End If
'CASO NEGATIVO, LIBERAÇÃO DOS CAMPOS E LIMPEZA PARA PREENCHIMENTO CORRETO
    If UserForm4.CBCADASTRO.Value = "COLABORADOR NOVO" Then
        
        UserForm4.txtnome2.Enabled = True
        UserForm4.txtnome2.Text = ""
        
        UserForm4.txtempresa2.Enabled = True
        UserForm4.txtempresa2.Text = ""
        UserForm4.txtempresa2.BackColor = &HFFFFFF 'BRANCO
        UserForm4.txtempresa2.ForeColor = &H8000000D 'AZUL
        
        UserForm4.txtfuncao2.Enabled = True
        UserForm4.txtfuncao2.Text = ""
        UserForm4.txtfuncao2.BackColor = &HFFFFFF 'BRANCO
        UserForm4.txtfuncao2.ForeColor = &H8000000D 'AZUL
        
    End If
'ALTERAÇÃO - SE SELECIONADO, AJUSTAR OS CAMPOS SINALIZADOS
    If UserForm4.CBCADASTRO.Value = "ATUALIZAÇÃO" Then
        
        UserForm4.txtnome2.Enabled = False
        UserForm4.txtnome2.Text = Sheets("dados").Range("H2").Value

        UserForm4.txtempresa2.Enabled = True
        UserForm4.txtempresa2.Text = Sheets("dados").Range("I2").Value
        UserForm4.txtempresa2.BackColor = &HFF& 'VERMELHO
        UserForm4.txtempresa2.ForeColor = &HFFFFFF  'BRANCO
        
        UserForm4.txtfuncao2.Enabled = True
        UserForm4.txtfuncao2.Text = Sheets("dados").Range("J2").Value
        UserForm4.txtfuncao2.BackColor = &HFF&  'VERMELHO
        UserForm4.txtfuncao2.ForeColor = &HFFFFFF   'BRANCO
        
    End If

End Sub

Private Sub cbnovocadastro2_Click()

'VALIDAÇÃO SE NÃO TEM CAMPO EM BRANCO
    If txtnome2.Value = "" _
    Or txtnome2.Value = " " _
    Or txtempresa2.Value = "" _
    Or txtempresa2.Value = " " _
    Or txtfuncao2.Value = "" _
    Or txtfuncao2.Value = " " _
    Or UserForm4.CBCADASTRO.Value = "" _
    Or UserForm4.CBCADASTRO.Value = " " Then
        'CASO ESTEJA VAZIO INFORMAR SOBRE O CAMPO PENDENTE
            MsgBox "Favor, Informar os campos em branco!", vbCritical, "Erro!"
            UserForm4.CBCADASTRO.SetFocus
'CASO NAO ESTEJA VAZIO CONTINUAR COM O CADASTRO
    Else
       ' Call novo_cadastro
    
    ' Gravar os valores em outra planilha
    Dim wbDestino As Workbook
    Dim wsDestino As Worksheet
    Set wbDestino = Workbooks.Open("H:\Hospital\Almoxarifado\RELATÓRIOS PFF\SENTINELA 1.0\BD_CADASTRO.xlsx") ' Abra a planilha de destino
    Set wsDestino = wbDestino.Sheets("BD_CADASTRO")
    
    wbDestino.Application.Visible = False
    
    Dim lastRow As Long
    lastRow = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row + 1
    
    wsDestino.Cells(lastRow, 1).Value = UserForm1.txtcpf.Value
    wsDestino.Cells(lastRow, 2).Value = UserForm4.txtnome2.Value
    wsDestino.Cells(lastRow, 3).Value = UserForm4.txtempresa2.Value
    wsDestino.Cells(lastRow, 4).Value = UserForm4.txtfuncao2.Value
    wsDestino.Cells(lastRow, 5).Value = UCase(Application.UserName)
    wsDestino.Cells(lastRow, 6).Value = Date
    wsDestino.Cells(lastRow, 7).Value = UserForm4.CBCADASTRO.Value
    
    ' Salvar e fechar a planilha de destino
    wbDestino.Save
    wbDestino.Close SaveChanges:=False
    
        
    ' Exibir mensagem de sucesso ou outras ações necessárias
    MsgBox "Dados gravados com sucesso!", vbInformation
      
       
  '_______________________
       
        Unload Me
        Call salvar
    End If

'CASO TENHA ERRO CONTINUE
    On Error Resume Next
        UserForm1.txtnome.Value = Sheets("dados").Range("H2").Value
    On Error Resume Next
        UserForm1.txtempresa.Value = Sheets("dados").Range("I2").Value
    On Error Resume Next
        UserForm1.txtfuncao.Value = Sheets("dados").Range("J2").Value
    On Error Resume Next
        UserForm1.txtstatus.Value = Sheets("dados").Range("K2").Value
    On Error Resume Next
        UserForm1.txtvalid.Value = Sheets("dados").Range("L2").Value

'FORMATAÇÃO PARA LIMPAR CAMPOS
    UserForm1.Frame1.BackColor = &HFFFFFF 'BRANCO
    UserForm1.txtnome.BackColor = &HFFFFFF 'BRANCO
    UserForm1.txtnome.ForeColor = &H8000000D 'AZUL
    UserForm1.TextBox1.ForeColor = &H8000000D 'AZUL
    UserForm1.txtempresa.BackColor = &HFFFFFF 'BRANCO
    UserForm1.txtempresa.ForeColor = &H8000000D 'AZUL
    UserForm1.TextBox2.ForeColor = &H8000000D 'AZUL
    UserForm1.txtfuncao.BackColor = &HFFFFFF 'BRANCO
    UserForm1.txtfuncao.ForeColor = &H8000000D 'AZUL
    UserForm1.TextBox3.ForeColor = &H8000000D 'AZUL
    UserForm1.Frame2.BackColor = &HFFFFFF 'BRANCO
    UserForm1.TextBox1.BackColor = &HFFFFFF 'BRANCO
    UserForm1.TextBox2.BackColor = &HFFFFFF 'BRANCO
    UserForm1.TextBox3.BackColor = &HFFFFFF 'BRANCO
    
'STATUS DA DISPENSA
    If UserForm1.txtstatus.Text = " " Then
        UserForm1.Frame1.BackColor = &HFFFFFF
    End If

    If UserForm1.txtstatus.Text = "LIBERADO" Then
        UserForm1.Frame1.BackColor = &HC000&
        UserForm1.txtstatus.ForeColor = &HFFFFFF
    End If

    If UserForm1.txtstatus.Text = "ANTES DO PRAZO" Then
        UserForm1.Frame1.BackColor = &HFF&
        UserForm1.txtstatus.ForeColor = &HFFFFFF
    End If
    
'FOCO NA JUSTIFICATIVA APOS CADASTRO
    UserForm1.cboxjustificativa.SetFocus
    
    Call ATUALIZAR_CADASTRO
    
    'CASO TENHA ERRO CONTINUE
    On Error Resume Next
        UserForm1.txtnome.Value = Sheets("dados").Range("H2").Value
    On Error Resume Next
        UserForm1.txtempresa.Value = Sheets("dados").Range("I2").Value
    On Error Resume Next
        UserForm1.txtfuncao.Value = Sheets("dados").Range("J2").Value

End Sub

Private Sub CommandButton1_Click()

'FECHAR FORMULARIO
    Unload Me

End Sub

Private Sub txtempresa2_Change()

'FORMATAÇÃO DE CAIXA AUTA
    txtempresa2 = UCase(txtempresa2.Value)
    
End Sub

Private Sub txtempresa2_Exit(ByVal Cancel As MSForms.ReturnBoolean)

'FORMANTANDO APOIS AJUSTE
    If UserForm4.txtempresa2.Value = "NÃO INFORMADO" Then
        UserForm4.txtempresa2.BackColor = &HFF&
        UserForm4.txtempresa2.ForeColor = &HFFFFFF
    Else
        UserForm4.txtempresa2.BackColor = &HFFFFFF
        UserForm4.txtempresa2.ForeColor = &H8000000D
    End If
    
End Sub

Private Sub txtfuncao2_Change()

'FORMATAÇÃO DE CAIXA AUTA
    txtfuncao2 = UCase(txtfuncao2.Value)
    
End Sub

Private Sub txtfuncao2_Exit(ByVal Cancel As MSForms.ReturnBoolean)

'FORMANTANDO APOIS AJUSTE
    If UserForm4.txtfuncao2.Value = "NÃO INFORMADO" Then
        UserForm4.txtfuncao2.BackColor = &HFF&
        UserForm4.txtfuncao2.ForeColor = &HFFFFFF
    Else
        UserForm4.txtfuncao2.BackColor = &HFFFFFF
        UserForm4.txtfuncao2.ForeColor = &H8000000D
    End If
    
End Sub

Private Sub txtnome2_Change()

'FORMATAÇÃO DE CAIXA AUTA
    txtnome2 = UCase(txtnome2.Value)

End Sub

Private Sub UserForm_Initialize()
  
'LISTA DE TIPO DE CADASTRO A SER EFETIVADO
    CBCADASTRO.RowSource = "DADOS!F3:F6"

'FORMATAÇÃO DE USUARIO EFETUANDO O CADASTRO
    UserForm4.Label14 = UCase(Application.UserName)

End Sub
