Attribute VB_Name = "Módulo1"
Dim inicio As Boolean

'MACRO DO RELOGIO
    Sub MeuRelogio()
        If inicio = True Then
            UserForm1.lbhora = Time
            Application.OnTime Now() + TimeValue("00:00:01"), "meurelogio"
        End If
            
    End Sub

'MACRO PARA INICIAR O RELOGIO
    Sub IniciarRelogio()

        inicio = True
    
        Call MeuRelogio

    End Sub

'MACRO DE DISPENSA DE INFORMAÇÕES NA PLANILHA
    Sub dispensar()

        Dim tabela3 As ListObject
        Dim t As Integer, ID As Integer

        Set tabela3 = Planilha3.ListObjects(1)
'            ID = Range("ID").Value

            t = tabela3.Range.Rows.Count
            
'INFORMAÇÕES A SEREM INFORMADAS
 '           tabela3.Range(t, 1).Value = ID
            tabela3.Range(t, 2).Value = UserForm1.txtcpf.Value
            tabela3.Range(t, 3).Value = UserForm1.txtnome.Value
            tabela3.Range(t, 4).Value = UserForm1.cboxjustificativa.Value
            tabela3.Range(t, 5).Value = UserForm1.txtobs.Value
            tabela3.Range(t, 6).Value = Date
            tabela3.Range(t, 7).Value = Time
            tabela3.Range(t, 12).Value = UCase(Application.UserName)
            tabela3.Range(t, 13).Value = UserForm1.txtquant.Value
            tabela3.Range(t, 14).Value = UserForm1.lb_loc
            
'INCLUSAO DE NOVA LINHA
            tabela3.ListRows.Add
 '           Range("ID").Value = ID + 1

'CASO TUDO OK INFORMANDO DE PROCESSO EFETUADO COM SUCESSO!!
            MsgBox "Dispensado!", vbInformation, "Informação"
            
    End Sub

'MACRO PARA SALVAR PLANILHA
    Sub salvar()
        
        
            ActiveWorkbook.Save
    End Sub

'CADASTRANDO NOVOS COLABORADORES
'    Sub novo_cadastro()
'
'            Dim tabela4 As ListObject
'            Dim k As Integer
'
'            Set tabela4 = Planilha4.ListObjects(1)
'
'                k = tabela4.Range.Rows.Count
'
'INFORMAÇÕES A SEREM CADASTADOS COMO NOVO COLABORADOR
'            tabela4.Range(k, 1).Value = UserForm1.txtcpf.Value
'            tabela4.Range(k, 2).Value = UserForm4.txtnome2.Value
'            tabela4.Range(k, 3).Value = UserForm4.txtempresa2.Value
'            tabela4.Range(k, 4).Value = UserForm4.txtfuncao2.Value
'            tabela4.Range(k, 5).Formula = UCase(Application.UserName)
'            tabela4.Range(k, 6).Formula = Date
'
'NOVA LINHA PARA O NOVO CADASTRO
'            tabela4.ListRows.Add
'
'CASO TUDO OK, INFORMATIVO DE PROCESSO EFETUADO COM SUCESSO!!
'        If UserForm4.CBCADASTRO.Value = "COLABORADOR NOVO" Then
'            MsgBox "Cadastrado com Sucesso!", vbInformation, "Informação"
'        Else
'            MsgBox "Atualizado com Sucesso!", vbInformation, "Informação"
'        End If
'
'    End Sub

'ABRIR FORMULARIO DA PLANILHA RELATORIO
 Sub chamar()
Attribute chamar.VB_ProcData.VB_Invoke_Func = " \n14"
 'Atalho do teclado: Ctrl+q
    Application.Visible = False
    UserForm1.Show
    UserForm1.txtcpf.SetFocus
End Sub

'FUNÇÃO PARA FORMULA DE CPF

'Função que valida CPF
Public Function lfValidaCPF(ByVal lNumCPF As String) As Boolean
    Application.Volatile
    
    Dim lMultiplicador  As Integer
    Dim lDv1            As Integer
    Dim lDv2            As Integer
    
    lMultiplicador = 2
    
    'Realiza o preenchimento dos zeros á esquerda
    lNumCPF = String(11 - Len(lNumCPF), "0") & lNumCPF
    
    'Realiza o cálculo do dividendo para o dv1 e o dv2
    For i = 9 To 1 Step -1
        lDv1 = (Mid(lNumCPF, i, 1) * lMultiplicador) + lDv1
        
        lDv2 = (Mid(lNumCPF, i, 1) * (lMultiplicador + 1)) + lDv2
        
        lMultiplicador = lMultiplicador + 1
    Next
    
    'Realiza o cálculo para chegar no primeiro dígio
    lDv1 = lDv1 Mod 11
    
    If lDv1 >= 2 Then
        lDv1 = 11 - lDv1
    Else
        lDv1 = 0
    End If
    
    'Realiza o cálculo para chegar no segundo dígido
    lDv2 = lDv2 + (lDv1 * 2)
    
    lDv2 = lDv2 Mod 11
    
    If lDv2 >= 2 Then
        lDv2 = 11 - lDv2
    Else
        lDv2 = 0
    End If
    
    'Realiza a validação e retorna na função
    If Right(lNumCPF, 2) = CStr(lDv1) & CStr(lDv2) Then
        lfValidaCPF = True
    Else
        lfValidaCPF = False
    End If
End Function

'EXPORTANDO RELATORIO COMPLETO
Sub REL_COMPLETO()

'SE TIVER ERRO INFORMAR O ERRO
    On Erro GoTo Erro
    
'SE NAO TIVER ERRO CONTINUE

    With Planilha3 'PLANILHA DE TRABALHO
    Dim lsCaminho As String 'CAMINHO PARA ABRIR/SALVAR RELATORIO
    Dim Linha As Integer 'LINHA QUE DEVE SER TRATADA
        
        lsCaminho = Planilha2.Range("S2").Cells 'AONDE BUSCAR O CAMINHO
        Linha = 2 'AONDE FICA A LINHA

'ABRINDO O CAMINHO
        Open lsCaminho For Output As 1

'VERIFICANDO SE NÃO ESTA VAZIA
        Do Until .Cells(Linha, 2) = ""
        
'SE PREENCHUDA, COLOCA NO RELATORIOS OS CAMPOS ABAIXO
    Print #1, .Cells(Linha, 2) & ";" & _
            .Cells(Linha, 3) & ";" & _
            .Cells(Linha, 4) & ";" & _
            .Cells(Linha, 5) & ";" & _
            .Cells(Linha, 6) & ";" & _
            .Cells(Linha, 7) & ";" & _
            .Cells(Linha, 8) & ";" & _
            .Cells(Linha, 9) & ";" & _
            .Cells(Linha, 10) & ";" & _
            .Cells(Linha, 11) & ";" & _
            .Cells(Linha, 12) & ";" & _
            .Cells(Linha, 13)

            Linha = Linha + 1
'REPETE O PROCESSO ATÉ ZERAR O CAMPO DA LINHA
        Loop

'QUANDO ZERADA INFORMA A FINALIZAÇÃO DO PROCESSO
           Resp = MsgBox("Concluido, Salvo em: " & lsCaminho, vbYesNo)
           If Resp - vbNo Then
           UserForm3.Hide
           Cancel = True
           Exit Sub
           End If
        Close 1
    End With
'SE NAO TIVER ERRO SEGUE NORMAL
sair:
    Exit Sub
    
'SE TIVER ERRO INFORMAR O TIPO DE ERRO
Erro:
    MsgBox "Erro na criptografia " & Err.Description
    GoTo sair

End Sub

'EXPORTANDO RELATORIO DAS JUSTIFICATIVAS
Sub REL_JUSTIFICATIVAS()

'SE TIVER ERRO INFORMAR O ERRO
    On Erro GoTo Erro
    
'SE NAO TIVER ERRO CONTINUE

    With Planilha3 'PLANILHA DE TRABALHO
    Dim lsCaminho1 As String 'CAMINHO PARA ABRIR/SALVAR RELATORIO
    Dim linha1 As Integer 'LINHA QUE DEVE SER TRATADA
        
        lsCaminho1 = Planilha2.Range("S2").Cells 'AONDE BUSCAR O CAMINHO
        linha1 = 2 'AONDE FICA A LINHA

'ABRINDO O CAMINHO
    Open lsCaminho1 For Output As 2

'VERIFICANDO SE NÃO ESTA VAZIA
    Do Until .Cells(linha1, 2) = ""
    
'SE PREENCHUDA, COLOCA NO RELATORIOS OS CAMPOS ABAIXO
    Print #2, .Cells(linha1, 2) & ";" & _
            .Cells(linha1, 3) & ";" & _
            .Cells(linha1, 4) & ";" & _
            .Cells(linha1, 6) & ";" & _
            .Cells(linha1, 7) & ";" & _
            .Cells(linha1, 8) & ";" & _
            .Cells(linha1, 9) & ";" & _
            .Cells(linha1, 11)
          
            linha1 = linha1 + 1
'REPETE O PROCESSO ATÉ ZERAR O CAMPO DA LINHA
    Loop

'QUANDO ZERADA INFORMA A FINALIZAÇÃO DO PROCESSO
            resp1 = MsgBox("Concluido, Salvo em: " & lsCaminho1, vbYesNo)
           If resp1 - vbNo Then
           UserForm3.Hide
           Cancel = True
           Exit Sub
           End If
        Close 2
    End With
    
'SE NAO TIVER ERRO SEGUE NORMAL
sair:
    Exit Sub
    
'SE TIVER ERRO INFORMAR O TIPO DE ERRO
Erro:
    MsgBox "Erro na criptografia " & Err.Description
    GoTo sair
    
End Sub

'EXPORTANDO RELATORIO DAS EMPRESAS
Sub REL_EMPRESAS()

'SE TIVER ERRO INFORMAR O ERRO
    On Erro GoTo Erro
    
'SE NAO TIVER ERRO CONTINUE

    With Planilha3 'PLANILHA DE TRABALHO
    Dim lsCaminho2 As String 'CAMINHO PARA ABRIR/SALVAR RELATORIO
    Dim linha2 As Integer 'LINHA QUE DEVE SER TRATADA
    
        lsCaminho2 = Planilha2.Range("S2").Cells 'AONDE BUSCAR O CAMINHO
        linha2 = 2 'AONDE FICA A LINHA

'ABRINDO O CAMINHO
    Open lsCaminho2 For Output As 3

'VERIFICANDO SE NÃO ESTA VAZIA
    Do Until .Cells(linha2, 2) = ""
    
'SE PREENCHUDA, COLOCA NO RELATORIOS OS CAMPOS ABAIXO
    Print #3, .Cells(linha2, 2) & ";" & _
            .Cells(linha2, 3) & ";" & _
            .Cells(linha2, 6) & ";" & _
            .Cells(linha2, 7) & ";" & _
            .Cells(linha2, 8) & ";" & _
            .Cells(linha2, 9) & ";" & _
            .Cells(linha2, 10) & ";" & _
            .Cells(linha2, 11)
          
        linha2 = linha2 + 1
'REPETE O PROCESSO ATÉ ZERAR O CAMPO DA LINHA
    Loop
    
'QUANDO ZERADA INFORMA A FINALIZAÇÃO DO PROCESSO
            resp2 = MsgBox("Concluido, Salvo em: " & lsCaminho2, vbYesNo)
           If resp2 - vbNo Then
           UserForm3.Hide
           Cancel = True
           Exit Sub
           End If
        Close 3
    End With
    
'SE NAO TIVER ERRO SEGUE NORMAL
sair:
    Exit Sub
    
'SE TIVER ERRO INFORMAR O TIPO DE ERRO
Erro:
    MsgBox "Erro na criptografia " & Err.Description
    GoTo sair
    
End Sub

'EXPORTANDO RELATORIO DAS EMPRESAS
Sub REL_CADASTRO()

'SE TIVER ERRO INFORMAR O ERRO
    On Erro GoTo Erro
        
'SE NAO TIVER ERRO CONTINUE

    With Planilha4 'PLANILHA DE TRABALHO
    Dim lsCaminho3 As String 'CAMINHO PARA ABRIR/SALVAR RELATORIO
    Dim linha3 As Integer 'LINHA QUE DEVE SER TRATADA
    
        lsCaminho3 = Planilha2.Range("S2").Cells 'AONDE BUSCAR O CAMINHO
        linha3 = 2 'AONDE FICA A LINHA
'ABRINDO O CAMINHO
    Open lsCaminho3 For Output As 4

'VERIFICANDO SE NÃO ESTA VAZIA
    Do Until .Cells(linha3, 1) = ""
    
'SE PREENCHUDA, COLOCA NO RELATORIOS OS CAMPOS ABAIXO
    Print #4, .Cells(linha3, 1) & ";" & _
            .Cells(linha3, 2) & ";" & _
            .Cells(linha3, 3) & ";" & _
            .Cells(linha3, 4)
          
        linha3 = linha3 + 1
'REPETE O PROCESSO ATÉ ZERAR O CAMPO DA LINHA
    Loop
    
'QUANDO ZERADA INFORMA A FINALIZAÇÃO DO PROCESSO
            resp3 = MsgBox("Concluido, Salvo em: " & lsCaminho3, vbYesNo)
           If resp3 - vbNo Then
           UserForm3.Hide
           Cancel = True
           Exit Sub
           End If
        Close 4
    End With
    
'SE NAO TIVER ERRO SEGUE NORMAL
sair:
    Exit Sub
    
'SE TIVER ERRO INFORMAR O TIPO DE ERRO
Erro:
    MsgBox "Erro na criptografia " & Err.Description
    GoTo sair
    
End Sub
