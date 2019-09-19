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
      Set pTbConta.xDb = pSys.xDb
   End If
   With pTbConta
      If pIDFORMAPGTO = gIDFORMAPGTO_Dinheiro Or pIDFORMAPGTO = gIDFORMAPGTO_Cheque Then
         Sql = "TPCONTA='D'"
      Else
         Sql = "TPCONTA='B'"
      End If
      Sql = Sql + " And  EVENDA=1"
      Sql = Sql + " And  IDLOJA=" & pSys.Propriedades("IDLOJA")
      
      
      If .Pesquisar(Ch_Where:=Sql) Then
         DefineContaLAN = True
      Else
         If ExibirPergunta("Não existe conta associada." & vbNewLine & vbNewLine & "Deseja contiuar?") = vbYes Then
            DefineContaLAN = True
         Else
            DefineContaLAN = False
         End If
      End If
   End With
End Function
Public Sub SetCustomEvent(ByRef pEvent As CalendarEvent, mvarSys As Object, _
                                    Optional pFLGCONFIRMADO, _
                                    Optional pFLGCANCELADO, _
                                    Optional pFLGREMARCADO, _
                                    Optional pIDATENDIMENTO, _
                                    Optional pSITATEND, _
                                    Optional pIDVENDA, _
                                    Optional pSITVENDA, _
                                    Optional pIDCLIENTE)
   
   Dim NgCal      As Object 'NG_Calendario
   Set NgCal = CriarObjeto("Calendario3R.NG_Calendario")  'New NG_Calendario
   Set NgCal.Sys = mvarSys
   
   '***************
   '* Propriedades Customizadas
   If Not IsMissing(pFLGCONFIRMADO) Then pEvent.CustomProperties("FLGCONFIRMADO") = pFLGCONFIRMADO
   If Not IsMissing(pFLGCANCELADO) Then pEvent.CustomProperties("FLGCANCELADO") = pFLGCANCELADO
   If Not IsMissing(pFLGREMARCADO) Then pEvent.CustomProperties("FLGREMARCADO") = pFLGREMARCADO
   If Not IsMissing(pIDATENDIMENTO) Then pEvent.CustomProperties("IDATENDIMENTO") = pIDATENDIMENTO
   If Not IsMissing(pSITATEND) Then pEvent.CustomProperties("SITATEND") = pSITATEND
   If Not IsMissing(pIDVENDA) Then pEvent.CustomProperties("IDVENDA") = pIDVENDA
   If Not IsMissing(pSITVENDA) Then pEvent.CustomProperties("SITVENDA") = pSITVENDA
   If Not IsMissing(pIDCLIENTE) Then pEvent.CustomProperties("IDCLIENTE") = pIDCLIENTE

   Call NgCal.SetCustomIcons(pEvent)
   Set NgCal = Nothing
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


