Private Sub Auto_Open()
 
    MsgBox "ATUALIZE a planilha antes de enviar os emails!", vbExclamation, "IMPORTANTE"
    
End Sub

Private Sub Lista_Faltas()

    Dim verifunc, func  As Double
    Dim i, j, k, w, h, contaEmail As Integer
    Dim diasinter, dsem, dsems, profi, pront As String
    Dim dias(1 To 1000) As String
    
    i = 1
    j = 1
    w = 2  'começa pela linha 2 da planilha
    h = 2
    m = 2
    Z = 0
    contaEmail = 0
    
    Application.ScreenUpdating = False
    
    '================== ordena a planilha ===================
    
    Classifica_Planilha
    
    '============== mensagem de confirmação de envio de mensagens de cobrança ===============
    
    h = MsgBox("Tabela atualizada ?" & Chr(10) & Chr(13) & Chr(13) & "Enviar As mensagens?", vbYesNoCancel, "Envio de mensagem de cobrança") ' (yes = 6) (no = 7) (cancelar = 2)
    'm = MsgBox("Enviar mensagem alternativa de FINAL DE PERÍODO ?" & Chr(10) & Chr(13) & Chr(13), vbYesNoCancel, "Final de Mês/Semestre") ' (yes = 6) (no = 7) (cancelar = 2)
    m = 7 'garante a mensagem 1 no envio do email
    
    '============== módulo principal ===============
    
    If h = 6 Then
        
        func = Range("e2") 'variável recebe matricula
        
        verifunc = func 'atribui matricula a variável de verificação
        
        profi = Range("d2") 'variável recebe nome
        
        Do Until func = 0 'percorre até o final
            
            Do Until verifunc <> func 'repete enquanto a matricula e variável de verificação são iguais (até encontrar matricula(func) diferente)
                
                pront = Range("a" & w) 'seleciona o prontuário
                
                dsem = Range("b" & w) 'seleciona o dia
                
                dsems = Range("c" & w) 'seleciona clínica
                
                dias(j) = pront & " - " & dsems & " " & dsem 'monta o vetor com os dados coletados
                
                w = w + 1
                j = j + 1
                
                verifunc = Range("e" & w) 'variável recebe a matricula da linha de baixo
                
            Loop 'sai do loop quando a matricula (func) é diferente
            
                j = j - 1
                
                If j = 1 Then '1 paciente
                    diasinter = " * " & dias(1) & "<br>"
                Else
                    If j = 2 Then '2 pacientes
                        diasinter = " * " & dias(1) & "<br>" & " * " & dias(2) & "<br>"
                    Else '>3 pacientes
                        diasinter = " * " & dias(1) & "<br>"
                        For Z = 2 To j
                            diasinter = diasinter & " * " & dias(Z) & "<br>"
                        Next Z
                    End If
                End If
                
        '=================== seleciona qual mensagem enviar ===================
        
            If m = 7 Then
                    GoSub mensagem01 'mensagem de mês normal
            End If
            
            If m = 6 Then
                    GoSub mensagem02 'mensagem de final de mês ou semestre
            End If
            
        '=================== repete o processo e reseta os contadores para serem reutilizados na próxima matrícula ===================
            
            contaEmail = contaEmail + 1
            func = Range("e" & w)
            profi = Range("d" & w)
            verifunc = func
            j = 1
            Z = 0
            diasinter = ""
        
        Loop
        
        MsgBox contaEmail & " mensagens enviadas com sucesso.", vbInformation, "Aviso" 'mensagem final de envio
    
    End If

Application.ScreenUpdating = True

Exit Sub

'=================== mensagem normal do mês ===================
mensagem01:
    
    ' VERIFICAÇÃO DE DESTINATÁRIOS EXCLUÍDOS
'    Select Case func
'        Case "20220835"
'            MsgBox "Matrícula " & func & " ignorada.", vbInformation, "Aviso"
'            Return ' Pula o envio de mensagem para essa matrícula e volta para as outras
'    End Select
    
    'Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error Resume Next
    Set OutApp = CreateObject("Outlook.Application")
    OutApp.Session.Logon
    Set OutMail = OutApp.CreateItem(0)
    
    On Error Resume Next
    
    With OutMail
        
        '.Display        'mostra o email antes de enviar
        
        .To = func & "@sarah.br"     'concatena matricula com o @ do email
        
        .sentonbehalfofname = "RegistrodeEstatisticas-CNCQ@sarah.br"     'seta o nome do destinatário
        
        .Subject = "Pacientes atendidos sem registro de estatística"
        
        '.Importance = 2 ' 2 = Alta prioridade
        
        line1 = IIf(Hour(Now()) < 12, "Bom dia!", "Boa tarde!")
        line2 = "Estamos fazendo a conferência da digitação de estatística e verificamos que "
        line3 = IIf(j = 1, "há 1 paciente atendido na sua agenda, mas sem registro de estatística.", "há " & j & " pacientes atendidos na sua agenda, mas sem registro de estatística.")
        line5 = "Quando a estatística atrasada ultrapassar o limite de dias do sistema (10 dias), favor nos enviar o LOCAL e as ATIVIDADES realizadas."
        line6 = "Nos casos onde NÃO tenha ocorrido o atendimento ao paciente, solicitar o CANCELAMENTO DA RECEPÇÃO no setor de Atendimento ao Público."
        line7 = diasinter
        line8 = "Obrigado,"
        line9 = "Controle de Qualidade"
        line10 = "Ramal 1679"
        
        'CORPO DO EMAIL EM HTML (MÊS NORMAL)
                .HTMLBody = _
                    "<html>" & vbNewLine & _
                            "<body style=font-size:11pt;font-family:Calibri> " & vbNewLine & _
                                line1 & "<br><br>" & _
                                line2 & line3 & "<br><br>" & _
                                line5 & "<br><br>" & _
                                line6 & "<br><br>" & _
                                line7 & "<br><br>" & _
                                line8 & "<br><br>" & _
                                line9 & "<br>" & _
                                line10 & "<br>" & _
                            "</body>" & vbNewLine & _
                    "</html>"
        
        .Send
        
        'reseta a sessão de email
        Set OutApp = Nothing
        Set OutMail = Nothing
        
    End With
    On Error GoTo 0
    
    Application.DisplayAlerts = True
    'Application.ScreenUpdating = True

Return

'=================== mensagem de final de mês ===================
mensagem02:

    ' VERIFICAÇÃO DE DESTINATÁRIOS EXCLUÍDOS
'    Select Case func
'        Case "17000853"
'            ' MsgBox "Matrícula " & func & " ignorada.", vbInformation, "Aviso" 'aviso de não envio
'            Return ' Pula este envio e volta para o loop principal
'    End Select

    'Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error Resume Next
    Set OutApp = CreateObject("Outlook.Application")
    OutApp.Session.Logon
    Set OutMail = OutApp.CreateItem(0)
    
    On Error Resume Next
    
    With OutMail
        
        '.Display        'mostra o email antes de enviar
        
        .To = func & "@sarah.br"     'concatena matricula com o @ do email
        
        .sentonbehalfofname = "RegistrodeEstatisticas-CNCQ@sarah.br"     'seta o nome do destinatário
        
        .Subject = "Estatística final de mês"
        
        .Importance = 2 ' 2 = Alta prioridade
        
        line1 = IIf(Hour(Now()) < 12, "Bom dia!", "Boa tarde!")
        line2 = "Estamos conferindo a digitação de estatística do MÊS DE " & ObterMes & " e verificamos que " '<< TROCA DE MÊS AUTOMÁTICA
        line3 = IIf(j = 1, "há 1 paciente atendido na sua agenda, mas sem registro de estatística.", "há " & j & " pacientes atendidos na sua agenda, mas sem registro de estatística.")
        line5 = "Quando a estatística atrasada ultrapassar o limite de dias do sistema (10 dias), favor nos enviar o LOCAL e as ATIVIDADES realizadas."
        line6 = "Nos casos onde NÃO tenha ocorrido o atendimento ao paciente, solicitar o CANCELAMENTO DA RECEPÇÃO no setor de Atendimento ao Público."
        line7 = "<b><span style='color: red;'>=> Pedimos, por gentileza, digitar suas estatísticas ainda pendentes até amanhã, dia 04/11/2026. <=</span></b>" ' '<< TROCAR DATA LIMITE
        line8 = diasinter
        line9 = "Obrigado,"
        line10 = "Controle de Qualidade"
        line11 = "Ramal 1679"
        
        'CORPO DO EMAIL EM HTML COM A MENSAGEM DE FINAL DE MÊS OU SEMESTRAL (line5)
                .HTMLBody = _
                    "<html>" & vbNewLine & _
                            "<body style=font-size:11pt;font-family:Calibri> " & vbNewLine & _
                                line1 & "<br><br>" & _
                                line2 & line3 & "<br><br>" & _
                                line5 & "<br><br>" & _
                                line6 & "<br><br>" & _
                                line7 & "<br><br>" & _
                                line8 & "<br><br>" & _
                                line9 & "<br><br>" & _
                                line10 & "<br>" & _
                                line11 & "<br>" & _
                            "</body>" & vbNewLine & _
                    "</html>"
                
        .Send
        
        'reseta a sessão de email
        Set OutApp = Nothing
        Set OutMail = Nothing
        
    End With
    On Error GoTo 0
    
    Application.DisplayAlerts = True
    'Application.ScreenUpdating = True

Return

End Sub

'================ Classifica os dados da Planilha "Conferência" =======================

Private Sub Classifica_Planilha()

    Sheets("Conferência").Select
    ActiveWorkbook.Worksheets("Conferência").ListObjects("Tabela_Consulta_de_BSB").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Conferência").ListObjects("Tabela_Consulta_de_BSB").Sort.SortFields.Add Key:=Range("Tabela_Consulta_de_BSB[NR_MATRICULA]"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Conferência").ListObjects("Tabela_Consulta_de_BSB").Sort.SortFields.Add Key:=Range("Tabela_Consulta_de_BSB[DATA]"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Conferência").ListObjects("Tabela_Consulta_de_BSB").Sort.SortFields.Add Key:=Range("Tabela_Consulta_de_BSB[CLÍNICAS]"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Conferência").ListObjects("Tabela_Consulta_de_BSB").Sort.SortFields.Add Key:=Range("Tabela_Consulta_de_BSB[NR_REGISTRO]"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Conferência").ListObjects("Tabela_Consulta_de_BSB").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

'================ complemento para mensagem de final de mês =====================

Private Function ObterMes() As String

    Dim MesAtual As Integer
    Dim NomeMes As String
    MesAtual = Month(Date)

    Select Case MesAtual
        Case 1: NomeMes = "DEZEMBRO"
        Case 2: NomeMes = "JANEIRO"
        Case 3: NomeMes = "FEVEREIRO"
        Case 4: NomeMes = "MARÇO"
        Case 5: NomeMes = "ABRIL"
        Case 6: NomeMes = "MAIO"
        Case 7: NomeMes = "JUNHO"
        Case 8: NomeMes = "JULHO"
        Case 9: NomeMes = "AGOSTO"
        Case 10: NomeMes = "SETEMBRO"
        Case 11: NomeMes = "OUTUBRO"
        Case 12: NomeMes = "NOVEMBRO"

    End Select

    ObterMes = NomeMes

End Function

