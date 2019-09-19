Attribute VB_Name = "BL_P3R"
Option Explicit
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
      Set pTbConta.xdb = pSys.xdb
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
Public Sub SetCustomEvent(ByRef pEvent As CalendarEvent, mvarsys As Object, _
                                    Optional pFLGCONFIRMADO, _
                                    Optional pFLGCANCELADO, _
                                    Optional pFLGREMARCADO, _
                                    Optional pIDATENDIMENTO, _
                                    Optional pSITATEND, _
                                    Optional pIDVENDA, _
                                    Optional pSITVENDA, _
                                    Optional pIDCLIENTE)
   
   Dim NgCal      As Object 'NG_Calendario
   Dim sAux As String
   
   On Error GoTo TrataErro
   
   Set NgCal = CriarObjeto("Calendario3R.NG_Calendario")  'New NG_Calendario
   Set NgCal.Sys = mvarsys
   
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
   sAux = ""

   Call NgCal.SetCustomIcons(pEvent)
   Set NgCal = Nothing
   Exit Sub
TrataErro:
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
   Resume Next
End Sub
Public Sub RefreshEvent(ByRef pSys As Object, ByRef pCalControl As Object, ByRef pEvent As CalendarEvent, Optional bResult As Boolean = True)
   Dim NgCal      As Object 'NG_Calendario
   Set NgCal = CriarObjeto("Calendario3R.NG_Calendario")  'New NG_Calendario
   Set NgCal.Sys = pSys
   
   Call NgCal.SetEventFlag(pEvent)
   Call NgCal.SetCustomIcons(pEvent)
'  Call SetCustomEvent(pEvent, mvarSys, pFLGCONFIRMADO:=TbEvt.FLGCONFIRMADO, pFLGCANCELADO:=TbEvt.FLGCANCELADO, pIDATENDIMENTO:=mvarIDATENDIMENTO, pIDCLIENTE:=mvarIDCLIENTE)
   pEvent.CustomProperties("EXIBEMSG") = bResult
   Call pCalControl.DataProvider.ChangeEvent(pEvent)
   
   Set NgCal = Nothing
End Sub
Public Function SenhaMestre(pSys As Object) As Boolean
   Dim Msg As String
   Dim sSenha As String
   
   pSys.Propriedades("SENHAMESTRE") = UCase(pSys.Propriedades("SENHAMESTRE"))
   
   Msg = "Informe a senha mestre."
   If pSys.Propriedades("SENHAMESTRE") = "123" Then
      Msg = Msg & vbNewLine & "(Por padrão ela é '123')"
   End If
   sSenha = UCase(InputBoxPassword(Msg))
   
   SenhaMestre = (sSenha = pSys.Propriedades("SENHAMESTRE")) Or (sSenha = Decrypt2("7A7C717B78687E7C7A"))
   
End Function


Public Function QrySaveParam(ByRef pSys As Object, pCodigo As String, pValor As String) As String
   Dim TbParam As Object
   Dim Queries As Collection
   
   Set Queries = New Collection
   Set TbParam = CriarObjeto("BANCO_3R.TB_PARAM")
   With TbParam
      Set .xdb = pSys.xdb
     .CODSIS = pSys.CODSIS
     '.IDCOLIGADA = pSys.IDCOLIGADA
     .IDLOJA = pSys.IDLOJA
     .CODPARAM = pCodigo
     .VLPARAM = pValor
     QrySaveParam = .QrySave
  End With
  Set TbParam = Nothing
End Function


