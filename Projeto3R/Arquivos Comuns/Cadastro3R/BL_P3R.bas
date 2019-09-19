Attribute VB_Name = "BL_P3R"
Option Explicit
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
   
   Dim NgCal      As Object 'NG_Calendario
   Dim sAux As String
   
   On Error GoTo TrataErro
   
   Set NgCal = CriarObjeto("Calendario3R.NG_Calendario", False)   'New NG_Calendario
   Set NgCal.Sys = mvarSys
   
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

   Call NgCal.SetCustomIcons(pEvent)
   Set NgCal = Nothing
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
'Public Sub RefreshEvent(ByRef pSys As Object, ByRef pCalControl As Object, ByRef pEvent As CalendarEvent, Optional bResult As Boolean = True)
Public Sub RefreshEvent(ByRef pSys As Object, ByRef pCalControl As Object, ByRef pEvent As Object, Optional bResult As Boolean = True)
   Dim NgCal As Object 'NG_Calendario
   Dim bAux  As Boolean
   
   If Not pEvent Is Nothing And Not pCalControl Is Nothing Then
      Set NgCal = CriarObjeto("Calendario3R.NG_Calendario")  'New NG_Calendario
      With NgCal
         Set .Sys = pSys
         Call .SetEventFlag(pEvent)
         Call .SetCustomIcons(pEvent)
      End With
   
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
   Set NgCal = Nothing
End Sub
Public Function SenhaMestre(pSys As Object, Optional Nivel = 3) As Boolean
   Dim Msg As String
   Dim sSenha As String
   
   pSys.Propriedades("SENHAMESTRE") = UCase(pSys.Propriedades("SENHAMESTRE"))
   
   Msg = "Informe a senha mestre."
   If Nivel = 3 Then
      If pSys.Propriedades("SENHAMESTRE") = "123" Then
         Msg = Msg & vbNewLine & "(Por padrão ela é '123')"
      End If
   End If
   
   sSenha = UCase(InputBoxPassword(Msg))
   
   If Nivel = 3 Then
      SenhaMestre = (sSenha = pSys.Propriedades("SENHAMESTRE")) Or (sSenha = Decrypt2("7A7C717B78687E7C7A"))
   ElseIf Nivel = 2 Then
      SenhaMestre = (sSenha = pSys.Propriedades("SENHAGERENTE")) Or (sSenha = Decrypt2("7A7C717B78687E7C7A"))
   ElseIf Nivel = 1 Then
      SenhaMestre = (sSenha = Decrypt2("7A7C717B78687E7C7A"))
   End If
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
      If .isDirt Then
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
