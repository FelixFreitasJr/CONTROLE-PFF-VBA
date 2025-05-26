VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Relat�rio"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6375
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
'AO ABRIR O FORMULARIO � ATUALIZADO A LISTA DE SELE��O PARA O FORMATO DO RELATORIO
    CB_FORMATO.RowSource = "DADOS!D1:D4"
        
End Sub

Private Sub CB_FORMATO_Change()
'O QUE FOR SELECIONADO NA LISTA � ENVIADO PARA A PLANILHA
    Sheets("dados").Range("R2").Value = CB_FORMATO
 
End Sub

Private Sub BT_Completo_Click()

'SE O BOT�O DO RELATORIO DE CADASTRO CLICADO
    If BT_Completo = True Then

'QUANDO CLICADO VAI ACONTECER DE:
    
'INFORMAR NA PLANILHA O NOME DO ARQUIVO
    Sheets("dados").Range("Q2").Value = "Completo"

'DESTACAR A COR DA LETRA QUE ESTA SELECIONADO
    BT_Completo.ForeColor = &H8000000D

'BLOQUEAR OUTROS BOT�ES PARA EVITAR ERRO
    BT_Justificativa.Enabled = False
    BT_.Enabled = False
    BT_Cadastro.Enabled = False

    Else
'QUANDO N�O CLICADO VAI ACONTECER DE:
    
'RETORNAR A COR ORIGINAL PARA SELE��O
    BT_Completo.ForeColor = &H8000&

'LIBERAR O USO DOS OUTROS BOT�ES
    BT_Justificativa.Enabled = True
    BT_.Enabled = True
    BT_Cadastro.Enabled = True

    End If
End Sub

Private Sub BT_Justificativa_Click()

'SE O BOT�O DO RELATORIO DE CADASTRO CLICADO
    If BT_Justificativa = True Then

'QUANDO CLICADO VAI ACONTECER DE:
    
'INFORMAR NA PLANILHA O NOME DO ARQUIVO
    Sheets("dados").Range("Q2").Value = "Justificativa"

'DESTACAR A COR DA LETRA QUE ESTA SELECIONADO
    BT_Justificativa.ForeColor = &H8000000D

'BLOQUEAR OUTROS BOT�ES PARA EVITAR ERRO
    BT_.Enabled = False
    BT_Completo.Enabled = False
    BT_Cadastro.Enabled = False

    Else
'QUANDO N�O CLICADO VAI ACONTECER DE:
    
'RETORNAR A COR ORIGINAL PARA SELE��O
    BT_Justificativa.ForeColor = &H8000& ' VERDE

'LIBERAR O USO DOS OUTROS BOT�ES
    BT_.Enabled = True
    BT_Completo.Enabled = True
    BT_Cadastro.Enabled = True

    End If
End Sub

Private Sub BT__Click() ' EMPRESA

'SE O BOT�O DO RELATORIO DE CADASTRO CLICADO
    If BT_ = True Then

'QUANDO CLICADO VAI ACONTECER DE:
    
'INFORMAR NA PLANILHA O NOME DO ARQUIVO
    Sheets("dados").Range("Q2").Value = "Empresas"

'DESTACAR A COR DA LETRA QUE ESTA SELECIONADO
    BT_.ForeColor = &H8000000D

'BLOQUEAR OUTROS BOT�ES PARA EVITAR ERRO
    BT_Justificativa.Enabled = False
    BT_Completo.Enabled = False
    BT_Cadastro.Enabled = False

    Else
'QUANDO N�O CLICADO VAI ACONTECER DE:
    
'RETORNAR A COR ORIGINAL PARA SELE��O
    BT_.ForeColor = &H8000&

'LIBERAR O USO DOS OUTROS BOT�ES
    BT_Justificativa.Enabled = True
    BT_Completo.Enabled = True
    BT_Cadastro.Enabled = True

    End If
End Sub

Private Sub BT_Cadastro_Click()

'SE O BOT�O DO RELATORIO DE CADASTRO CLICADO
    If BT_Cadastro = True Then

'QUANDO CLICADO VAI ACONTECER DE:
    
'INFORMAR NA PLANILHA O NOME DO ARQUIVO
    Sheets("dados").Range("Q2").Value = "Cadastro"

'DESTACAR A COR DA LETRA QUE ESTA SELECIONADO
    BT_Cadastro.ForeColor = &H8000000D 'AZUL

'BLOQUEAR OUTROS BOT�ES PARA EVITAR ERRO
    BT_Justificativa.Enabled = False
    BT_Completo.Enabled = False
    BT_.Enabled = False

    Else
'QUANDO N�O CLICADO VAI ACONTECER DE:
    
'RETORNAR A COR ORIGINAL PARA SELE��O
    BT_Cadastro.ForeColor = &H8000& 'VERDE

'LIBERAR O USO DOS OUTROS BOT�ES
    BT_Justificativa.Enabled = True
    BT_Completo.Enabled = True
    BT_.Enabled = True

    End If
End Sub

Private Sub BTLIMPAR2_Click() ' LIMPAR CAMPOS

'LIMPAR CAMPOS E VOLTAR A CORES ORIGINAIS
    BT_Completo.Enabled = True
    BT_Completo.ForeColor = &H8000& 'VERDE

    BT_Justificativa.Enabled = True
    BT_Justificativa.ForeColor = &H8000& ' VERDE

    BT_.Enabled = True
    BT_.ForeColor = &H8000& 'VERDE

    BT_Cadastro.Enabled = True
    BT_Cadastro.ForeColor = &H8000& 'VERDE

    CB_FORMATO.Value = ""

End Sub

Private Sub cbfechar_rel_Click() 'FECHANDO FORMULARIO

Unload Me

End Sub

Private Sub bt_dispensar_rel_Click()

'SE O BOT�O SELECIONADO
    If BT_Completo = True Then
' CHAMA A MACRO PARA O RELATORIO SELECIONADO
    Call REL_COMPLETO

'SE O BOT�O SELECIONADO
    ElseIf BT_Justificativa = True Then
' CHAMA A MACRO PARA O RELATORIO SELECIONADO
    Call REL_JUSTIFICATIVAS

'SE O BOT�O SELECIONADO
    ElseIf BT_ = True Then 'EMPRESA
' CHAMA A MACRO PARA O RELATORIO SELECIONADO
    Call REL_EMPRESAS

'SE O BOT�O SELECIONADO
    ElseIf BT_Cadastro = True Then
' CHAMA A MACRO PARA O RELATORIO SELECIONADO
    Call REL_CADASTRO

    End If
    
'LIMPAR CAMPOS E VOLTAR A CORES ORIGINAIS
    BT_Completo.Enabled = True
    BT_Completo.ForeColor = &H8000& 'VERDE

    BT_Justificativa.Enabled = True
    BT_Justificativa.ForeColor = &H8000& ' VERDE

    BT_.Enabled = True
    BT_.ForeColor = &H8000& 'VERDE

    BT_Cadastro.Enabled = True
    BT_Cadastro.ForeColor = &H8000& 'VERDE

    CB_FORMATO.Value = ""
    
End Sub
