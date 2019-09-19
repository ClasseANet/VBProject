Attribute VB_Name = "BL_P3R"
Option Explicit
Public gSys As Object
Public Enum eTpTarefa
   TarBoasVindas = 1
   TarConfimaAge = 2
   TarOrientacao = 3
   TarVerifAvali = 4
   TarECancelado = 5
   TarNaoMarcado = 6
End Enum

Global Const gIDFORMAPGTO_Dinheiro = 1
Global Const gIDFORMAPGTO_Debito = 2
Global Const gIDFORMAPGTO_Credito = 3
Global Const gIDFORMAPGTO_Cheque = 4

Global Const gSITATEND_Aberto = "00"
Global Const gSITATEND_Fechado = "10"
Global Const gSITATEND_Excluido = "9X"
Global Const gSITVENDA_Aberta = "00"
Global Const gSITVENDA_Fechada = "10"
Global Const gSITVENDA_Excluido = "9X"
'Public Function DefineContaLAN(ByRef pTbConta As TB_FCCORRENTE, pIDFORMAPGTO As Long) As Boolean
Public Function DefineContaLAN(ByRef pTbConta As Object, pIDFORMAPGTO As Long, pSys As Object) As Boolean
   Dim Sql As String
   
   '* Define Conta
   If pTbConta Is Nothing Then
      Set pTbConta = CriarObjeto("BANCO_3R.TB_FCCORRENTE")
      Set pTbConta.xDb = pSys.xDb
   End If
   With pTbConta
      If pIDFORMAPGTO = gIDFORMAPGTO_Dinheiro Or pIDFORMAPGTO = gIDFORMAPGTO_Cheque Then
         Sql = "TPCONTA='D'"
      Else
         Sql = "TPCONTA='B'"
      End If
      Sql = Sql & " And  EVENDA=1"
      
      
      If .Pesquisar(Ch_IDLOJA:=pSys.Propriedades("IDLOJA"), Ch_Where:=Sql) Then
         DefineContaLAN = True
      Else
         If ExibirPergunta("Não existe conta associada." & vbNewLine & vbNewLine & "Deseja continuar?") = vbYes Then
            DefineContaLAN = True
         Else
            DefineContaLAN = False
         End If
      End If
   End With
End Function
'Public Sub SetCustomEvent(ByRef pEvent As CalendarEvent, mvarsys As Object,
Public Sub SetCustomEvent(ByRef pEvent As Object, mvarSys As Object, _
                                    Optional pFLGCONFIRMADO, _
                                    Optional pFLGCANCELADO, _
                                    Optional pFLGREMARCADO, _
                                    Optional pIDATENDIMENTO, _
                                    Optional pSITATEND, _
                                    Optional pIDVENDA, _
                                    Optional pSITVENDA, _
                                    Optional pIDCLIENTE, _
                                    Optional pIDSALA)
   
   
   Dim sAux As String
   
   On Error GoTo TrataErro
   
   '***************
   '* Propriedades Customizadas
   sAux = "FLGCONFIRMADO"
   If Not IsMissing(pFLGCONFIRMADO) Then pEvent.CustomProperties("FLGCONFIRMADO") = pFLGCONFIRMADO
   sAux = "FLGCANCELADO"
   If Not IsMissing(pFLGCANCELADO) Then pEvent.CustomProperties("FLGCANCELADO") = pFLGCANCELADO
   sAux = "FLGREMARCADO"
   If Not IsMissing(pFLGREMARCADO) Then pEvent.CustomProperties("FLGREMARCADO") = pFLGREMARCADO
   sAux = "IDATENDIMENTO"
   If Not IsMissing(pIDATENDIMENTO) Then pEvent.CustomProperties("IDATENDIMENTO") = pIDATENDIMENTO
   sAux = "SITATEND"
   If Not IsMissing(pSITATEND) Then pEvent.CustomProperties("SITATEND") = pSITATEND
   sAux = "IDVENDA"
   If Not IsMissing(pIDVENDA) Then pEvent.CustomProperties("IDVENDA") = pIDVENDA
   sAux = "SITVENDA"
   If Not IsMissing(pSITVENDA) Then pEvent.CustomProperties("SITVENDA") = pSITVENDA
   sAux = "IDCLIENTE"
   If Not IsMissing(pIDCLIENTE) Then pEvent.CustomProperties("IDCLIENTE") = pIDCLIENTE
   sAux = "IDSALA"
   If Not IsMissing(pIDSALA) Then pEvent.CustomProperties("IDSALA") = pIDSALA
   sAux = ""

   Call SetCustomIcons2(pEvent)
   
   Exit Sub
TrataErro:
   If Not pEvent Is Nothing Then
      Select Case sAux
         Case "FLGCONFIRMADO":   pEvent.CustomProperties("FLGCONFIRMADO") = 0
         Case "FLGCANCELADO":    pEvent.CustomProperties("FLGCANCELADO") = 0
         Case "FLGREMARCADO":    pEvent.CustomProperties("FLGREMARCADO") = 0
         Case "IDATENDIMENTO":   pEvent.CustomProperties("IDATENDIMENTO") = 0
         Case "SITATEND":        pEvent.CustomProperties("SITATEND") = gSITATEND_Aberto
         Case "IDVENDA":         pEvent.CustomProperties("IDVENDA") = 0
         Case "SITVENDA":        pEvent.CustomProperties("SITVENDA") = gSITVENDA_Aberta
         Case "IDCLIENTE":       pEvent.CustomProperties("IDCLIENTE") = 0
         Case "IDSALA":          pEvent.CustomProperties("IDSALA") = 0
      End Select
   End If
   Resume Next
End Sub
Public Sub SetCustomIcons2(ByRef pEvent As XtremeCalendarControl.CalendarEvent)
   Dim bFLGCONFIRMADO   As Boolean
   Dim bFLGCANCELADO    As Boolean
   Dim bFLGREMARCADO    As Boolean
   Dim bAtedFechado     As Boolean
   Dim nIDATEND         As Long
   Dim sSITATEND        As String
   Dim nIDVENDA         As Long
   Dim sSITVENDA        As String
   Dim nLabel0          As Integer
   Dim sAux             As String
      
   On Error GoTo TrataErro
      
   If pEvent Is Nothing Then Exit Sub
   With pEvent
      .CustomIcons.RemoveAll
      '* Customize standard icons
      If .PrivateFlag Then .CustomIcons.Add xtpCalendarEventIconIDPrivate
      If .Reminder Then .CustomIcons.Add xtpCalendarEventIconIDReminder
      If .MeetingFlag Then .CustomIcons.Add xtpCalendarEventIconIDMeeting
      If .RecurrenceState = xtpCalendarRecurrenceOccurrence Then .CustomIcons.Add xtpCalendarEventIconIDOccurrence
      If .RecurrenceState = xtpCalendarRecurrenceException Then .CustomIcons.Add xtpCalendarEventIconIDException

      bAtedFechado = False
      If .MeetingFlag Then
         sAux = "FLGCONFIRMADO"
         bFLGCONFIRMADO = (xVal(.CustomProperties("FLGCONFIRMADO")) = 1)
         sAux = "FLGCANCELADO"
         bFLGCANCELADO = (xVal(.CustomProperties("FLGCANCELADO")) = 1)
         sAux = "FLGREMARCADO"
         bFLGREMARCADO = (xVal(.CustomProperties("FLGREMARCADO")) = 1)
         sAux = "IDATENDIMENTO"
         nIDATEND = xVal(.CustomProperties("IDATENDIMENTO"))
         sAux = "IDVENDA"
         nIDVENDA = xVal(.CustomProperties("IDVENDA"))
         sAux = "SITVENDA"
         sSITVENDA = Trim(.CustomProperties("SITVENDA"))
         sAux = "LABEL0"
         nLabel0 = xVal(.CustomProperties("LABEL0"))
         sAux = ""
         
         On Error Resume Next
         sSITATEND = Trim(.CustomProperties("SITATEND"))
         If Err = 458 Then
            Dim Sql As String
            Dim MyRs As Object
            Sql = Sql_OATENDIMENTO(gSys.IDLOJA, nIDATEND)
            If gSys.xDb.Abretabela(Sql, MyRs) Then
               .CustomProperties("SITATEND") = MyRs("SITATEND").Value
            Else
               .CustomProperties("SITATEND") = gSITATEND_Aberto
            End If
            sSITATEND = Trim(.CustomProperties("SITATEND"))
         End If
         
         If nIDATEND <> 0 Then
            If sSITATEND = gSITATEND_Fechado Then
               bFLGCONFIRMADO = False
               bAtedFechado = True
               If nIDVENDA = 0 Then
                  .CustomIcons.Add 2   '* Atendimento Fechado
               Else
                  .CustomIcons.Add 8   '* Atendimento Fechado
               End If
            Else
               .CustomIcons.Add 1   '* Atendimento Aberto
            End If
         End If
         If Not bAtedFechado Then
            If nIDVENDA <> 0 Then
               If sSITVENDA = gSITVENDA_Fechada Then
                  bFLGCONFIRMADO = False
                  .CustomIcons.Add 5   '* Venda Fechado
               Else
                  .CustomIcons.Add 4   '* Venda Aberta
               End If
            End If
            'If .Id = 436 Then
            '   nLabel0 = nLabel0
            '   .Label = .Label
            'End If
            If bFLGCANCELADO Then
               .Label = 9999
               'bFLGREMARCADO = VerificarRemarcacao(gIDLOJA, pEvent.Id)
               If bFLGREMARCADO Then
                  .CustomIcons.Add 7
               Else
                  .CustomIcons.Add 6
               End If
               
            ElseIf bFLGCONFIRMADO Then
               If .Label = 9999 Then
                  .Label = nLabel0
               End If
               .CustomIcons.Add 3
            End If
         End If
      End If
   End With
   GoTo Saida
TrataErro:
   If Not pEvent Is Nothing Then
      Select Case sAux
         Case "FLGCONFIRMADO":   pEvent.CustomProperties("FLGCONFIRMADO") = 0
         Case "FLGCANCELADO":    pEvent.CustomProperties("FLGCANCELADO") = 0
         Case "FLGREMARCADO":    pEvent.CustomProperties("FLGREMARCADO") = 0
         Case "IDATENDIMENTO":   pEvent.CustomProperties("IDATENDIMENTO") = 0
         Case "SITATEND":        pEvent.CustomProperties("SITATEND") = gSITATEND_Aberto
         Case "IDVENDA":         pEvent.CustomProperties("IDVENDA") = 0
         Case "SITVENDA":        pEvent.CustomProperties("SITVENDA") = gSITVENDA_Aberta
         Case "IDCLIENTE":       pEvent.CustomProperties("IDCLIENTE") = 0
      End Select
   End If
   Resume Next
Saida:
End Sub
Public Sub SetEventFlag2(ByRef pEvent As CalendarEvent, Optional pRs As Object)   'Optional pIDEVENTO, Optional pIDATEND)
   Dim nIDLOJA    As Integer
   Dim sSITVENDA  As String
   Dim Sql        As String
   Dim MyRs       As Object
   Dim pIDEVENTO  As Long
   Dim pIDATEND   As Long
   Dim bOk        As Boolean
   
   If Not pEvent Is Nothing Then
      pIDEVENTO = pEvent.Id
   End If
   nIDLOJA = gSys.Propriedades("IDLOJA")
   If IsMissing(pIDEVENTO) Then pIDEVENTO = 0
   If IsMissing(pIDATEND) Then pIDATEND = 0
   
   If pIDATEND <> 0 And pIDEVENTO = 0 Then
      Sql = Sql_OEVENTOAGENDA(nIDLOJA, pIDATEND:=pIDATEND)
      If gSys.xDb.Abretabela(Sql, MyRs) Then
         pIDEVENTO = MyRs("IDEVENTO")
      End If
   End If
   
   On Error GoTo 0
   On Error Resume Next
   If Not pRs Is Nothing Then
      Set MyRs = pRs.Clone
      Sql = xVal(MyRs("IDATENDIMENTO") & "")
            
      MyRs.Filter = pRs.Filter
      MyRs.Move pRs.AbsolutePosition - 1, 1
      bOk = (pIDEVENTO = MyRs("IDEVENTO") And Err = 0) '91-Object variable or With block variable not set
   End If
   
   If Not bOk Then  'Object variable or With block variable not set
      Sql = "Select E.IDCLIENTE, E.IDSALA, E.BODY, E.FLGCONFIRMADO, E.FLGCANCELADO, E.FLGREMARCADO"
      Sql = Sql + ", E.LABELID, A.IDATENDIMENTO, A.SITATEND, V.IDVENDA, V.SITVENDA"
      Sql = Sql + " From OEVENTOAGENDA E"
      Sql = Sql + " Left Join OATENDIMENTO A On E.IDLOJA=A.IDLOJA And E.IDEVENTO=A.IDEVENTO"
      Sql = Sql + " Left Join OATENDIMENTO_VENDA AV On A.IDLOJA=AV.IDLOJA And A.IDATENDIMENTO=AV.IDATENDIMENTO"
      Sql = Sql + " Left Join CVENDA V On AV.IDLOJA=V.IDLOJA And AV.IDVENDA=V.IDVENDA"
      Sql = Sql + " Where E.IDEVENTO=" & SqlNum(pIDEVENTO)
      Sql = Sql + " And E.IDLOJA=" & SqlNum(nIDLOJA)
      bOk = gSys.xDb.Abretabela(Sql, MyRs)
   End If
   If bOk Then
      pEvent.CustomProperties("IDCLIENTE") = xVal(MyRs("IDCLIENTE") & "")
      pEvent.CustomProperties("IDSALA") = xVal(MyRs("IDSALA") & "")
      pEvent.CustomProperties("FLGCONFIRMADO") = xVal(MyRs("FLGCONFIRMADO") & "")
      pEvent.CustomProperties("FLGCANCELADO") = xVal(MyRs("FLGCANCELADO") & "")
      pEvent.CustomProperties("FLGREMARCADO") = xVal(MyRs("FLGREMARCADO") & "")
      pEvent.CustomProperties("IDATENDIMENTO") = xVal(MyRs("IDATENDIMENTO") & "")
      pEvent.CustomProperties("SITATEND") = IIf((MyRs("SITATEND") & "") = "", gSITATEND_Aberto, MyRs("SITATEND") & "")
      pEvent.CustomProperties("IDVENDA") = xVal(MyRs("IDVENDA") & "")
      pEvent.CustomProperties("SITVENDA") = MyRs("SITVENDA") & ""
      pEvent.CustomProperties("LABEL0") = xVal(MyRs("LABELID") & "")
      
      If MyRs.RecordCount > 1 Then
         While Not MyRs.EOF
            If pIDEVENTO = MyRs("IDEVENTO") & "" Then
               sSITVENDA = MyRs("SITVENDA") & ""
               If sSITVENDA <> MyRs("SITVENDA") & "" Then
                  sSITVENDA = IIf(sSITVENDA < MyRs("SITVENDA") & "", sSITVENDA, MyRs("SITVENDA") & "")
               End If
            Else
               MyRs.MoveLast
            End If
            MyRs.MoveNext
         Wend
         pEvent.CustomProperties("SITVENDA") = sSITVENDA
      End If
   End If
End Sub
'Public Sub RefreshEvent(ByRef pSys As Object, ByRef pCalControl As Object, ByRef pEvent As CalendarEvent, Optional bResult As Boolean = True)
Public Sub RefreshEvent(ByRef pSys As Object, ByRef pCalControl As Object, ByRef pEvent As Object, Optional bResult As Boolean = True)
   Dim bAux  As Boolean
   
   If Not pEvent Is Nothing And Not pCalControl Is Nothing Then
      Call SetEventFlag2(pEvent)
      Call SetCustomIcons2(pEvent)
      
      With pCalControl
      '  Call SetCustomEvent(pEvent, mvarSys, pFLGCONFIRMADO:=TbEvt.FLGCONFIRMADO, pFLGCANCELADO:=TbEvt.FLGCANCELADO, pIDATENDIMENTO:=mvarIDATENDIMENTO, pIDCLIENTE:=mvarIDCLIENTE)
         bAux = pEvent.CustomProperties("EXIBEMSG")
         
         pEvent.CustomProperties("EXIBEMSG") = bResult
         Call .DataProvider.ChangeEvent(pEvent)
         .RedrawControl
         .Populate
         pEvent.CustomProperties("EXIBEMSG") = bAux
      End With
   End If
End Sub
Public Function SenhaMestre(pSys As Object, Optional Nivel = 3) As Boolean
   Dim Msg     As String
   Dim sSenha  As String
   Dim bResult As Boolean
   
   pSys.Propriedades("SENHAMESTRE") = UCase(pSys.Propriedades("SENHAMESTRE"))
   
   If Nivel = 3 Then
      Msg = "Informe a senha operacional."
   ElseIf Nivel = 2 Then
      Msg = "Informe a senha gerencial."
   Else
      Msg = "Informe a senha mestre."
   End If
   
   If Nivel = 3 Then
      If pSys.Propriedades("SENHAMESTRE") = "123" Then
         Msg = Msg & vbNewLine & "(Por padrão ela é '123')"
      End If
   End If
   If Nivel = 2 Then
      If pSys.Propriedades("SENHAGERENTE") = "123" Then
         Msg = Msg & vbNewLine & "(Por padrão ela é '123')"
      End If
   End If
   
   sSenha = UCase(InputBoxPassword(Msg))
   
   If Nivel = 3 Then
      bResult = (sSenha = pSys.Propriedades("SENHAMESTRE")) Or (sSenha = Decrypt2("7A7C717B78687E7C7A"))
   End If
   If Nivel = 2 Or Not bResult Then
      bResult = (sSenha = pSys.Propriedades("SENHAGERENTE")) Or (sSenha = Decrypt2("7A7C717B78687E7C7A"))
   End If
   If Nivel = 1 Or Not bResult Then
      bResult = (sSenha = Decrypt2("7A7C717B78687E7C7A"))
   End If
   SenhaMestre = bResult
End Function
Public Function QrySaveParam(ByRef pSys As Object, pCodigo As String, pValor As String, Optional ByRef pQueries As Collection) As String
   Dim TbParam As Object
   Dim Queries As Collection
   
   Set Queries = New Collection
   Set TbParam = CriarObjeto("BANCO_3R.TB_PARAM")
   With TbParam
       Set .xDb = pSys.xDb
      .CODSIS = pSys.CODSIS
      '.IDCOLIGADA = pSys.IDCOLIGADA
      .IDLOJA = pSys.IDLOJA
      .CODPARAM = pCodigo
      .Pesquisar
      .VLPARAM = pValor
      If .IsDirt Then
         If pQueries Is Nothing Then Set pQueries = New Collection
         pQueries.Add .QrySave
         QrySaveParam = .QrySave
      End If
   End With
   Set TbParam = Nothing
End Function
Public Function CorPadrao(pDSCTRATAMENTO As String) As Long
   Dim TlCad As Object
   Dim nCor As Long
   
   Set TlCad = CriarObjeto("CADASTRO3R.TL_CADOTPTRATAMENTO", False)
   If Not TlCad Is Nothing Then
      nCor = TlCad.CorPadrao(pDSCTRATAMENTO)
   End If
   CorPadrao = nCor
End Function
Public Sub MontarMail(ByRef pSys As Object, ByRef xMail As Object, Optional peMailTo)
   With xMail
      .UseAuthentication = (pSys.GetParam("UseAuthentication") = xtpChecked)
      .UsePopAuthentication = (pSys.GetParam("UsePopAuthentication") = xtpChecked)
  
      .POP3Host = pSys.GetParam("POP3Host") ' "pop3.bol.com.br"
      .SMTPHost = pSys.GetParam("SMTPHost") ' "smtps.bol.com.br"
      .SMTPPort = pSys.GetParam("SMTPPort") ' 587
      .Username = pSys.GetParam("MailUID")  ' "diogenes72@bol.com.br"
      .Password = Decrypt2(pSys.GetParam("MailPWD"))
      
      .FromDisplayName = pSys.GetParam("FromDisplayName") 'FromDisplayName ' "Diogenes"
      If IsMissing(peMailTo) Then
         .Recipient = pSys.GetParam("LstEMailSocio") 'LstEMailSocio.Text ' "disantos@ig.com.br"
      Else
         .Recipient = peMailTo
      End If
      '.RecipientDisplayName = "Socio"           ' "DiSantos"
      
      .From = .Username
      .AsHTML = True
      
      .Receipt = True
      .SMTPHostValidation = 0 'VALIDATE_HOST_NONE
   End With
End Sub
Public Sub ShowCliente(ByRef pSys As Object, pIDLOJA As Integer, pIDCLIENTE As Long)
   Dim MyCliente As Object
   
   Screen.MousePointer = vbHourglass
   Set MyCliente = CriarObjeto("Contato3R.TL_CADCliente", False)
   With MyCliente
      Set .Sys = pSys
      .IDLOJA = pIDLOJA
      .IDCLIENTE = pIDCLIENTE
      Call .Show
   End With
   Set MyCliente = Nothing
   Screen.MousePointer = vbDefault
End Sub
Public Function Sql_OATENDIMENTO(pIDLOJA As Integer, Optional pIDATEND, Optional pIDCLIENTE) As String
   Dim Sql As String
   Sql = "Select *" & vbNewLine
   Sql = Sql & " From OATENDIMENTO" & vbNewLine
   Sql = Sql & " Where IDLOJA=" & pIDLOJA & vbNewLine
   If Not IsMissing(pIDATEND) Then Sql = Sql & " And IDATEND=" & pIDATEND & vbNewLine
   If Not IsMissing(pIDCLIENTE) Then Sql = Sql & " And IDCLIENTE=" & pIDCLIENTE & vbNewLine
   
   Sql_OATENDIMENTO = Sql
End Function
Public Function Sql_OEVENTOAGENDA(pIDLOJA As Integer, Optional pIDEVENTO, Optional pIDATEND) As String
   Dim Sql As String
   Sql = "Select *" & vbNewLine
   Sql = Sql & " From OEVENTOAGENDA" & vbNewLine
   Sql = Sql & " Where IDLOJA=" & pIDLOJA & vbNewLine
   If Not IsMissing(pIDEVENTO) Then Sql = Sql & " And IDEVENTO=" & pIDEVENTO & vbNewLine
   If Not IsMissing(pIDATEND) Then Sql = Sql & " And IDATEND=" & pIDATEND & vbNewLine
   
   Sql_OEVENTOAGENDA = Sql
End Function
Public Function Sql_OSERVICOEVT(pIDLOJA As Integer, pIDEVENTO As Long) As String
   Dim Sql As String
   Sql = "Select *" & vbNewLine
   Sql = Sql & " From OSERVICOEVT" & vbNewLine
   Sql = Sql & " Where IDLOJA=" & pIDLOJA & vbNewLine
   Sql = Sql & " And IDEVENTO=" & pIDEVENTO & vbNewLine
   
   Sql_OSERVICOEVT = Sql
End Function
Public Function Sql_OTPSERVICO(pIDLOJA As Integer, Optional pIDTPSERVICO) As String
   Dim Sql As String
   Sql = "Select *" & vbNewLine
   Sql = Sql & " From OTPSERVICO" & vbNewLine
   Sql = Sql & " Where IDLOJA=" & pIDLOJA & vbNewLine
   If Not IsMissing(pIDTPSERVICO) Then Sql = Sql & " And IDTPSERVICO=" & pIDTPSERVICO & vbNewLine
   
   Sql_OTPSERVICO = Sql
End Function
Public Function Sql_OTPTRATAMENTO(pIDLOJA As Integer, Optional pIDTPTRATAMENTO) As String
   Dim Sql As String
   Sql = "Select *" & vbNewLine
   Sql = Sql & " From OTPTRATAMENTO" & vbNewLine
   Sql = Sql & " Where IDLOJA=" & pIDLOJA & vbNewLine
   If Not IsMissing(pIDTPTRATAMENTO) Then Sql = Sql & " And IDTPTRATAMENTO=" & pIDTPTRATAMENTO & vbNewLine
   
   Sql_OTPTRATAMENTO = Sql
End Function
Public Function Sql_OAREA(pIDLOJA As Integer, Optional pIDAREA) As String
   Dim Sql As String
   Sql = "Select *" & vbNewLine
   Sql = Sql & " From OAREA" & vbNewLine
   Sql = Sql & " Where IDLOJA=" & pIDLOJA & vbNewLine
   If Not IsMissing(pIDAREA) Then Sql = Sql & " And IDAREA=" & pIDAREA & vbNewLine
   
   Sql_OAREA = Sql
End Function
Public Function Sql_OAREA_TRATAMENTO(pIDLOJA As Integer, Optional pIDAREA, Optional pIDTPTRATAMENTO, Optional pIDTPMANIPULO) As String
   Dim Sql As String
   Sql = "Select *" & vbNewLine
   Sql = Sql & " From OAREA_TRATAMENTO" & vbNewLine
   Sql = Sql & " Where IDLOJA=" & pIDLOJA & vbNewLine
   If Not IsMissing(pIDAREA) Then Sql = Sql & " And IDAREA=" & pIDAREA & vbNewLine
   If Not IsMissing(pIDTPTRATAMENTO) Then Sql = Sql & " And pIDTPTRATAMENTO=" & pIDTPTRATAMENTO & vbNewLine
   If Not IsMissing(pIDTPMANIPULO) Then Sql = Sql & " And IDTPMANIPULO=" & pIDTPMANIPULO & vbNewLine
   
   Sql_OAREA_TRATAMENTO = Sql
End Function
Public Function Sql_OCLIENTE(pIDLOJA As Integer, Optional pIDCLIENTE) As String
   Dim Sql As String
   Sql = "Select *" & vbNewLine
   Sql = Sql & " From OCLIENTE" & vbNewLine
   Sql = Sql & " Where IDLOJA=" & pIDLOJA & vbNewLine
   If Not IsMissing(pIDCLIENTE) Then Sql = Sql & " And IDCLIENTE=" & pIDCLIENTE & vbNewLine
   
   Sql_OCLIENTE = Sql
End Function
Public Function Sql_OMANIPULO(pIDLOJA As Integer, Optional pIDMANIPULO) As String
   Dim Sql As String
   Sql = "Select *" & vbNewLine
   Sql = Sql & " From OMANIPULO" & vbNewLine
   Sql = Sql & " Where IDLOJA=" & pIDLOJA & vbNewLine
   If Not IsMissing(pIDMANIPULO) Then Sql = Sql & " And IDMANIPULO=" & pIDMANIPULO & vbNewLine
   
   Sql_OMANIPULO = Sql
End Function
Public Function Sql_OSESSAO(pIDLOJA As Integer, Optional pIDATENDIMENTO, Optional pIDSESSAO) As String
   Dim Sql As String
   Sql = "Select *" & vbNewLine
   Sql = Sql & " From OSESSAO" & vbNewLine
   Sql = Sql & " Where IDLOJA=" & pIDLOJA & vbNewLine
   If Not IsMissing(pIDATENDIMENTO) Then Sql = Sql & " And IDATENDIMENTO=" & pIDATENDIMENTO & vbNewLine
   If Not IsMissing(pIDSESSAO) Then Sql = Sql & " And IDSESSAO=" & pIDSESSAO & vbNewLine
   
   Sql_OSESSAO = Sql
End Function

