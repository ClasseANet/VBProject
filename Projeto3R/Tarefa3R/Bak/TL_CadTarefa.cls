VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CadTarefa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe    As FrmCadTarefa
Attribute mvarMe.VB_VarHelpID = -1
Dim nLenIDVenda   As Integer
Dim bVendaFechada As Boolean
Dim dToday        As Date

Private mvarSys            As Object
Private mvarIDLOJA         As Integer
Private mvarIDATENDIMENTO  As Long
Private mvarIDCLIENTE      As Long
Private mvarIDVENDA        As Long
Private mvarCalControl     As CalendarControl
Private mvarCalEvent       As CalendarEvent
Public Property Set CalEvent(ByRef vData As Object)
   Set mvarCalEvent = vData
End Property
Public Property Get CalEvent() As Object
    Set CalEvent = mvarCalEvent
End Property
Public Property Set CalControl(ByVal vData As Object)
   Set mvarCalControl = vData
   On Error Resume Next
   mvarCalControl.DataProvider.ClearCache
End Property
Public Property Get CalControl() As Object
    Set CalControl = mvarCalControl
End Property
Public Property Let IDVENDA(ByVal vData As Long)
    mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Long
   IDVENDA = mvarIDVENDA
End Property
Public Property Let IDCLIENTE(ByVal vData As Long)
    mvarIDCLIENTE = vData
End Property
Public Property Get IDCLIENTE() As Long
   IDCLIENTE = mvarIDCLIENTE
End Property
Public Property Let IDATENDIMENTO(ByVal vData As Long)
    mvarIDATENDIMENTO = vData
End Property
Public Property Get IDATENDIMENTO() As Long
   IDATENDIMENTO = mvarIDATENDIMENTO
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
    mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   
   If mvarIDLOJA = 0 Then mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
   nLenIDVenda = mvarSys.Propriedades("LENIDVENDA")
   On Error Resume Next
   dToday = CDate(Format(mvarSys.xDb.sysdate, "DD/MM/YYYY"))
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmCadTarefa
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
End Sub
Public Sub Show(Optional pMode As FormShowConstants = vbModal)
'   If mvarIDLOJA = 0 Then
'      Call ExibirAviso("Código da Loja Inválido! (""" & mvarIDLOJA & """)")
'      Exit Sub
'   End If
   
   mvarMe.Show pMode
End Sub
Private Sub mvarMe_Activate()
   'Call PopulaTela
   
   If GetTag(mvarMe, "1VEZ", "1") = "1" Then
      Call SetTag(mvarMe, "1VEZ", "0")
      With mvarMe.GrdPagamento
         If .RowCount <= 1 Then
            .SetFocus
            Call .SetCurCell(.RowCount, 1)
         End If
      End With
'      If xVal(mvarMe.TxtVLVENDA.Text) = 0 Then
'         mvarMe.CmdOk.Enabled = False
'         ExibirAviso ("Venda com valor igual a zero.")
'         'Unload mvarMe
'       End If
   End If
End Sub
Private Sub mvarMe_CmdCancelClick()
   Unload mvarMe
End Sub

Private Sub mvarMe_CmdChaveClick()
   If InputBoxPassword("Informe a Senha.") = mvarSys.Propriedades("SENHAMESTRE") Then
      DoEvents
      bVendaFechada = False
      Call SetTag(mvarMe, "1VEZ", "1")
      'Call LimparTela
      Call PopulaTela
      Call SetTag(mvarMe, "1VEZ", "0")
      mvarMe.CmdFechar.Enabled = True
      mvarMe.CmdExcluir.Visible = True
   End If
End Sub

Private Sub mvarMe_CmdExcluirClick()
   Dim sMsg As String
   Dim Queries As Collection
   Dim Sql As String
   Dim nRecibo As Long
   Dim MyRs As Object
   
   Dim TbVenda As Object 'TB_CVENDA
   Dim TbItem As Object 'TB_CITENSVENDA
   Dim TbPgto As Object 'TB_CPGTOSVENDA
   Dim TbAten As Object 'TB_OATENDIMENTO_VENDA
   Dim TbReci As Object 'TB_FRECIBO
   
   sMsg = "Esta operação irá excluir definitivamente a venda '" & mvarMe.TxtIDVENDA.Text & "'"
   sMsg = sMsg & vbNewLine & vbNewLine
   sMsg = sMsg & "Deseja continuar?"
   If ExibirPergunta(sMsg, , False) = vbYes Then
      If InputBoxPassword("Informe a senha.") = mvarSys.Propriedades("SENHAMESTRE") Then
         Set Queries = New Collection
         
         Set TbVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
         With TbVenda
            Set .xDb = mvarSys.xDb
            If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA) Then
               Set TbItem = CriarObjeto("BANCO_3R.TB_CITENSVENDA")
               Set TbPgto = CriarObjeto("BANCO_3R.TB_CPGTOSVENDA")
               Set TbAten = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_VENDA")
               Set TbReci = CriarObjeto("BANCO_3R.TB_FRECIBO")
               Set TbItem.xDb = mvarSys.xDb
               Set TbPgto.xDb = mvarSys.xDb
               Set TbAten.xDb = mvarSys.xDb
               Set TbReci.xDb = mvarSys.xDb
               
               Queries.Add TbItem.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA)
               Queries.Add TbPgto.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA)
               Queries.Add TbAten.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA)
               Queries.Add .QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA)
               
               nRecibo = xVal(Trim(Mid(mvarMe.CmdRecibo.Caption, Len("Recibo "))))
               If TbReci.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDRECIBO:=nRecibo, Ch_SERIE:=1) Then
                  Sql = "Select Max(IDRECIBO)"
                  Sql = Sql & " From FRECIBO "
                  Sql = Sql & " Where IDLOTE=0 "
                  Sql = Sql & " And SERIE=1"
                  Sql = Sql & " And IDLOJA=" & SqlNum(mvarIDLOJA)
                  If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
                     If nRecibo = xVal(MyRs(0) & "") Then
                        Queries.Add TbReci.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDRECIBO:=nRecibo, Ch_SERIE:=1)
                     Else
                     
                        TbReci.Status = 2
                        Queries.Add TbReci.QryUpdate
                     End If
                  End If
               End If
               Call SalvarEvento(True)
               If mvarSys.xDb.Executa(Queries) Then
                  Call ExibirResultado(mvarSys, True)
                  Unload mvarMe
               Else
                  Call SalvarEvento
                  Call ExibirResultado(mvarSys, False)
               End If
            End If
         End With
      Else
         Call ExibirAviso("Esta senha não permite a operação.")
      End If
   End If
End Sub

Private Sub mvarMe_CmdFecharClick()
   If FecharVenda Then
      Unload mvarMe
   End If
End Sub
Private Function FecharVenda() As Boolean
   Dim Queries As Collection
   Dim TbVenda As Object 'TB_CVENDA
   Dim TbConta As Object 'TB_FCCORRENTE
   Dim TbFLAN  As Object 'TB_FLAN
   Dim TbPgto  As Object 'TB_CPGTOSVENDA
   Dim TbItens As Object 'TB_CITENSVENDA
   Dim TbEstP  As Object 'TB_SESTPROD
   Dim TbMovE  As Object 'TB_SMOVEST
   
   Dim Sql     As String
   Dim bOk     As Boolean
   Dim i       As Integer
   Dim iParcela As Integer
   Dim tParcela As Integer
   Dim nIDEST  As Long
   Dim MyRs    As Object
   Dim MyRsAux As Object
   Dim bNovo   As Boolean
   
   If Not VerificaFechamento() Then Exit Function
   If Not SalvarVenda() Then Exit Function
   
   Set Queries = New Collection
   
   '******************
   '* Fechar Venda
   '******************
   Set TbVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
   With TbVenda
      .xDb = mvarSys.xDb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA) Then
         .SITVENDA = gSITVENDA_Fechada
         Queries.Add .QrySave
      End If
   End With
   
   '******************
   '* Lançar Pagamento
   '******************
   Set TbConta = CriarObjeto("BANCO_3R.TB_FCCORRENTE")
   TbConta.xDb = mvarSys.xDb
   
   Set TbFLAN = CriarObjeto("BANCO_3R.TB_FLAN")
   TbFLAN.xDb = mvarSys.xDb
      
   Set TbPgto = CriarObjeto("BANCO_3R.TB_CPGTOSVENDA")
   Set TbPgto.xDb = mvarSys.xDb
   If TbPgto.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA, Ch_ORDERBY:="IDPGTO") Then
      While Not TbPgto.Rs.EOF
         TbPgto.Popula
   
         '*************
         '* Define Conta
         If Not DefineContaLAN(TbConta, TbPgto.IDFORMAPGTO, mvarSys) Then GoTo Saida
                              
'         Sql = "TPDOC='CV'"
'         Sql = Sql & " And IDFOR=1"
'         Sql = Sql & " And TPTRANSA=" & SqlStr(IIf(TbPgto.IDFORMAPGTO = 4, "Q", "C"))
'         Sql = Sql & " And IDLOJA = " & SqlNum(mvarIDLOJA)
'         Sql = Sql & " And IDCONTA=" & TbConta.IDCONTA
'         Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
         
         tParcela = IIf(TbPgto.IDCONDPGTO > 1, TbPgto.IDCONDPGTO - 1, 1)
         For iParcela = 1 To tParcela
            Sql = " IDLOJA = " & SqlNum(mvarIDLOJA)
            Sql = Sql & " And IDVENDA = " & SqlNum(mvarIDVENDA)
            Sql = Sql & " And IDPGTO = " & SqlNum(TbPgto.IDPGTO)
            Sql = Sql & " And NPARCELA = " & SqlNum(iParcela)
                                          
            Set TbFLAN = CriarObjeto("BANCO_3R.TB_FLAN")
            With TbFLAN
               .xDb = mvarSys.xDb
               bNovo = Not TbFLAN.Pesquisar(Ch_Where:=Sql)
               If bNovo Then
                  .CODSIT = ""
                  .DTCADASTRO = mvarSys.xDb.sysdate(3)
               End If
               
               .IDLOJA = mvarIDLOJA
               .IDCONTA = TbConta.IDCONTA
               
               .IDVENDA = mvarIDVENDA
               .IDPGTO = TbPgto.IDPGTO
               
               .NDOC = StrZero(TbPgto.IDVENDA, mvarSys.Propriedades("LENIDVENDA"))
               .Valor = TbPgto.VLPGTO - IIf(TbPgto.IDFORMAPGTO = 1, TbVenda.VLTROCO, 0)
               
               .DTEMISSAO = TbVenda.DTVENDA  '* XXX
               .DTVENCIMENTO = .DTEMISSAO
               .DTBAIXA = .DTEMISSAO
                          
               .IDFAVORECIDO = 1
               .FAVORECIDO = "Cliente"
               .TPFAVORECIDO = "C"
               .TPLAN = "C"
               .TPTRANSA = IIf(TbPgto.IDFORMAPGTO = 4, "Q", "C")
               .TPDOC = "CV"
               
               .IDDESP = 1
               .IDSUBDESP = 0
   '            .SALDO = SqlNum(0, True)
   'XXX           .SALDO = CalcularSaldo
               
               '.DTCONTAB = ""
               '.DTBAIXA = ""
               Call DefineParcela(TbFLAN, TbVenda.DTVENDA, TbPgto.IDFORMAPGTO, TbPgto.IDCONDPGTO, TbPgto.VLPGTO, iParcela)
               If .isDirt Then
                  Queries.Add .QrySave
               End If
            End With
         Next
         TbPgto.Rs.MoveNext
      Wend
   End If
      
   '******************
   '* Baixar estoque
   '******************
   Sql = "Select V.IDLOJA, V.IDVENDA, V.DTVENDA, V.IDCLIENTE, I.IDITEM"
   Sql = Sql & ", I.IDPROD, I.QTDVENDA, I.UNIDCONTROLE"
   Sql = Sql & " From CVENDA V Join CITENSVENDA I On V.IDLOJA=I.IDLOJA And V.IDVENDA=I.IDVENDA "
   Sql = Sql & " Join SPRODUTO P On I.IDPROD=P.IDPROD"
   Sql = Sql & " Where P.SITPROD='A'"
   Sql = Sql & " And P.EVENDA=1"
   Sql = Sql & " And P.ESERVICO=0"
   Sql = Sql & " And V.IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And V.IDVENDA=" & mvarIDVENDA
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      Sql = "Select L.IDEST"
      Sql = Sql & " From SEST_LOJA L Join SESTOQUE E On L.IDEST=E.IDEST  "
      Sql = Sql & " Where E.ATIVO=1"
      Sql = Sql & " And L.IDLOJA=" & mvarIDLOJA
      If mvarSys.xDb.AbreTabela(Sql) Then
         nIDEST = Val(mvarSys.xDb.RsAux("IDEST") & "")
      End If
      If nIDEST > 0 Then
         While Not MyRs.EOF
            Set TbEstP = CriarObjeto("BANCO_3R.TB_SESTPROD")
            Set TbEstP.xDb = mvarSys.xDb
            
            Call TbEstP.Pesquisar(Ch_IDEST:=nIDEST, Ch_IDPROD:=MyRs("IDPROD"))
               
            TbEstP.IDEST = nIDEST
            TbEstP.IDPROD = MyRs("IDPROD")
            If TbEstP.UNIDCONTROLE = "" Then
               TbEstP.UNIDCONTROLE = MyRs("UNIDCONTROLE")
            ElseIf TbEstP.UNIDCONTROLE <> MyRs("UNIDCONTROLE") Then
               If ExibirPergunta("Unidade de medida está diferente do Estoque." & vbNewLine & vbNewLine & "Deseja contiuar?", pDefaultYes:=False) = vbNo Then
                  GoTo Saida
               End If
            End If
            TbEstP.SLDATUAL = TbEstP.SLDATUAL - xVal(MyRs("QTDVENDA"))
            TbEstP.SLDDISPONIVEL = TbEstP.SLDDISPONIVEL - xVal(MyRs("QTDVENDA"))
            
            If TbEstP.SLDATUAL < 0 Or TbEstP.SLDDISPONIVEL < 0 Then
               If ExibirPergunta("Saldo de produto negativo." & vbNewLine & vbNewLine & "Deseja contiuar?", , pDefaultYes:=False) = vbNo Then
                  GoTo Saida
               End If
            End If
            
            
               
            Set TbMovE = CriarObjeto("BANCO_3R.TB_SMOVEST")
            With TbMovE
               Set .xDb = mvarSys.xDb
               Sql = "TPDOC='CV'"
               'Sql = Sql & " And NUMDOC=" & StrZero(MyRs("IDVENDA") & "", mvarSys.Propriedades("LENIDVENDA"))
               Sql = Sql & " And IDVENDA=" & MyRs("IDVENDA") & ""
               Sql = Sql & " And ITEMDOC=" & MyRs("IDITEM") & ""
               Sql = Sql & " And IDFOR=1"
               Sql = Sql & " And IDPROD=" & MyRs("IDPROD") & ""
               'Sql = Sql & " And IDFOR=" & MyRs("IDCLIENTE") & ""
               If .Pesquisar(Ch_Where:=Sql) Then
                  TbEstP.SLDATUAL = TbEstP.SLDATUAL + xVal(MyRs("QTDVENDA"))
                  TbEstP.SLDDISPONIVEL = TbEstP.SLDDISPONIVEL + xVal(MyRs("QTDVENDA"))
               End If
               Queries.Add TbEstP.QrySave
               
               If TbMovE.Rs.RecordCount > 1 Then
                  If mvarSys.xDb.AbreTabela("Select -1* Sum(QTDITEM)as [QTD] From SMOVEST Where IDMOV<>" & .IDMOV & " And " & Sql, MyRsAux) Then
                     TbEstP.SLDATUAL = TbEstP.SLDATUAL + xVal(MyRsAux("QTD"))
                     TbEstP.SLDDISPONIVEL = TbEstP.SLDDISPONIVEL + xVal(MyRsAux("QTD"))
                     Queries.Add TbEstP.QrySave
                  End If
                  Queries.Add "Delete From SMOVEST Where IDMOV<>" & .IDMOV & " And " & Sql
               End If
               .IDEST = nIDEST
               .IDPROD = MyRs("IDPROD") & ""
               .IDLOJA = mvarIDLOJA
               .DTMOV = MyRs("DTVENDA") & ""
               .QTDITEM = -1 * xVal(MyRs("QTDVENDA") & "")
               .UNIDCONTROLE = MyRs("UNIDCONTROLE") & ""
               .TPDOC = "CV"
               .IDVENDA = MyRs("IDVENDA")
               .NUMDOC = StrZero(MyRs("IDVENDA") & "", mvarSys.Propriedades("LENIDVENDA"))
               .IDFOR = 1 'MyRs("IDCLIENTE") & ""
               .ITEMDOC = MyRs("IDITEM") & ""
               .TPLAN = "D"
               .TPTRANSA = "D"
'XXX              .SALDO = CalcularSaldoEst
'               .SALDO = SqlNum(0, True)
               Queries.Add .QrySave
            End With
            
            MyRs.MoveNext
         Wend
      End If
      
   End If

   If Queries.Count > 0 Then
      FecharVenda = mvarSys.xDb.Executa(Queries)
   End If
   If Not FecharVenda Then
      Call ExibirInformacao("Erro ao fechar venda.")
   End If
   
Saida:
   Set Queries = Nothing
   Set TbVenda = Nothing
   Set TbConta = Nothing
   Set TbFLAN = Nothing
   Set TbPgto = Nothing
   Set TbItens = Nothing
   Set TbEstP = Nothing
   Set TbMovE = Nothing
End Function
Private Sub DefineParcela(ByRef pTbFLAN As Object, pDTVENDA As String, pIDFORMAPGTO As Integer, pIDCONDPGTO As Integer, pVLPGTO As Double, pNParcela As Integer)
   Dim i As Integer
   
   With pTbFLAN
   
      Select Case pIDFORMAPGTO
         Case 1 '* Dinheiro
            .DTVENCIMENTO = pDTVENDA
            .NPARCELA = 1
            .tParcela = 1
   
         Case 2 '* Débito
            .DTVENCIMENTO = DateAdd("d", 1, CDate(pDTVENDA))
            .NPARCELA = 1
            .tParcela = 1
      
         Case 3 '* Crédito
            'For i = 1 To pIDCONDPGTO - 1 '* Prazo, 1, 2 ou 3 vezes
               .DTVENCIMENTO = DateAdd("d", 30 * pNParcela, CDate(pDTVENDA))
               .NPARCELA = pNParcela
               .tParcela = pIDCONDPGTO - 1
               
               If pNParcela = pIDCONDPGTO - 1 Then
                  .Valor = xVal(ValBr(pVLPGTO - (xVal(ValBr(pVLPGTO / .tParcela, 2)) * (.tParcela - 1))))
               Else
                  .Valor = xVal(ValBr(pVLPGTO / .tParcela, 2))
               End If
               
            'Next
         
         Case 4 '* Cheque
            .TPTRANSA = "Q"
            If pIDCONDPGTO = 1 Then '* À vista
               .DTVENCIMENTO = DateAdd("d", 1, CDate(pDTVENDA))
               .NPARCELA = i
               .tParcela = pIDCONDPGTO - 1
               
            Else
               'For i = 1 To pIDCONDPGTO - 1 '* Prazo, 1, 2 ou 3 vezes
                  .DTVENCIMENTO = DateAdd("d", 30 * pNParcela, CDate(pDTVENDA))
                  .NPARCELA = pNParcela
                  .tParcela = pIDCONDPGTO - 1
               'Next
            End If
            
      End Select
   End With
End Sub
Public Sub CalcularSaldo(pIDLOJA As Integer, pIDCONTA As Integer, pDTBAIXA As String)
   Dim Sql As String

   Sql = "Select Top(10) IDLOJA , IDCONTA, IDLAN, CONVERT(datetime,ISNULL(DTBAIXA, GETDATE()),105) AS DTORDER, SALDO"
   Sql = Sql & "From FLAN"
   Sql = Sql & "Where IDLOJA = " & SqlNum(pIDLOJA)
   Sql = Sql & "And IDCONTA= " & SqlNum(pIDCONTA)
   Sql = Sql & "and DTEMISSAO<=" & SqlDate(pDTBAIXA) '2010-10-23 09:30:00.000'"
   Sql = Sql & "Order By DTORDER DESC, DTEMISSAO, IDLAN"
End Sub
Private Function VerificaFechamento() As Boolean
   If Not VerificaVenda() Then Exit Function
   
   With mvarMe.GrdPagamento
      If .RowCount <= 1 Then
         If xVal(UnFormat(mvarMe.TxtVLTROCO.Text)) = 0 Then
            If vbYes = ExibirPergunta("O valor de pagamento é zero." & vbNewLine & vbNewLine & "Deseja continuar?") Then
               If Not InputBoxPassword("Informe a senha.") = mvarSys.Propriedades("SENHAMESTRE") Then
                  Call ExibirInformacao("Pagamento inválido.")
                  .SetFocus
                  Exit Function
               End If
            Else
               Call ExibirInformacao("Pagamento inválido.")
               .SetFocus
               Exit Function
            End If
         Else
            Call ExibirInformacao("Pagamento inválido.")
            .SetFocus
            Exit Function
         End If
      Else
         If Val(.CellValue(1, "IDFORMAPGTO")) = 0 Then
            Call ExibirInformacao("Forma de pagamento inválida.")
            Call .SetCurCell(1, "IDFORMAPGTO")
            .SetFocus
            Exit Function
         End If
         If Val(.CellValue(1, "IDCONDPGTO")) = 0 Then
            Call ExibirInformacao("Condição de pagamento inválida.")
            Call .SetCurCell(1, "IDCONDPGTO")
            .SetFocus
            Exit Function
         End If
         If xVal(.CellValue(1, "VLPGTO")) = 0 Then
            Call ExibirInformacao("Valor de pagamento inválido.")
            Call .SetCurCell(1, "VLPGTO")
            .SetFocus
            Exit Function
         End If
      End If
      If ValorReal(mvarMe.TxtVLTROCO.Text) < 0 Then
         Call ExibirInformacao("Valores inválidos")
         If mvarMe.GrdPagamento.RowCount > 1 Then
            .SetFocus
            Call .SetCurCell(.RowCount - 1, "VLPGTO")
         Else
            mvarMe.TxtVLPGTO.SetFocus
         End If
         Exit Function
      End If
   End With
   
   VerificaFechamento = True
End Function
Private Sub mvarMe_CmdIDCLIENTEClick()
   Dim MyContato As Object
   
   Set MyContato = CriarObjeto("Contato3R.TL_CADContato")
   With MyContato
      Set .Sys = mvarSys
      .IDCLIENTE = mvarIDCLIENTE
      Call .Show
   End With
   Set MyContato = Nothing
End Sub
Private Sub mvarMe_CmdLovCliClick()
   Dim MyContato As Object
   
   Set MyContato = CriarObjeto("Contato3R.TL_NovoLov")
   With MyContato
      Set .Sys = mvarSys
      .Filtro = Trim(mvarMe.TxtNOME.Text)
      Call .ShowClientes
      If .IdSelecionado <> 0 Then
         Call PopulaCliente(.IdSelecionado)
      End If
   End With
   Set MyContato = Nothing
End Sub

Private Sub mvarMe_CmdOkClick()
   If Not VerificaVenda() Then Exit Sub
   If SalvarVenda Then
      Unload mvarMe
   End If
End Sub
Private Function VerificaVenda() As Boolean
   Dim i As Long
   Dim bDin As Boolean
   'If Trim(mvarMe.TxtTEL1.Text) = "" Then
   '   Call ExibirInformacao("Informe o telefone do cliente.")
   '   mvarMe.TxtTEL1.SetFocus
   '   Exit Function
   'End If
   If Trim(mvarMe.TxtNOME.Text) = "" Then
      Call ExibirInformacao("Informe o nome do cliente.")
      mvarMe.TxtNOME.SetFocus
      Exit Function
   End If
   If GetTag(mvarMe.TxtNOME, "IDCLIENTE", "0") = "0" Then
      Call ExibirInformacao("Cliente sem identificação." & vbNewLine & "Favor verificar cadastro.")
      mvarMe.TxtTEL1.SetFocus
      Exit Function
   End If
   
   If CDate(Format(mvarMe.CmbDTVENDA.Value, "DD/MM/YYYY")) > dToday Then
      Call ExibirInformacao("Data da venda não pode ser maior que hoje." & vbNewLine & "Favor verificar data.")
      If mvarMe.CmbDTVENDA.Enabled Then
         mvarMe.CmbDTVENDA.SetFocus
      Else
         mvarMe.GrdVenda.SetFocus
      End If
      Exit Function
   End If
   
   
   With mvarMe.GrdVenda
      If .RowCount <= 1 Then
         Call ExibirInformacao("Itens de venda inválidos.")
         .SetFocus
         Exit Function
      Else
         If xVal(.CellValue(1, "Total")) = 0 Then
            Call ExibirInformacao("Valor do item inválido.")
            Call .SetCurCell(1, "QTDVENDA")
            .SetFocus
            Exit Function
         End If
         If xVal(.CellValue(1, "IDPROD")) = 0 Then
            
            Call ExibirInformacao("Produto inválido.")
            Call .SetCurCell(1, "IDPROD")
            .SetFocus
            Exit Function
         End If
      End If
   End With
   
   bDin = False
   If xVal(mvarMe.TxtVLTROCO) > 0 Then
      For i = 1 To mvarMe.GrdPagamento.RowCount - 1
         bDin = bDin Or (mvarMe.GrdPagamento.CellValue(i, "IDFORMAPGTO") = 1)
         If bDin Then Exit For
      Next
      If Not bDin Then
         Call ExibirInformacao("Valor de troco inválido." & vbNewLine & "Só poderá haver troco se o pagamento for em 'Dinheiro'.")
         Exit Function
      End If
   End If
   

   VerificaVenda = True
End Function
Private Function SalvarVenda() As Boolean
   Dim i          As Integer
   Dim iParcela   As Integer
   Dim tParcela   As Integer
   Dim cQueries   As Collection
   Dim bResult    As Boolean
   Dim bAchou     As Boolean
   Dim Sql        As String
   Dim sIn        As String
   Dim TbVenda    As Object 'TB_CVENDA
   Dim TBAtend    As Object 'TB_OATENDIMENTO_VENDA
   Dim TbItens    As Object 'TB_CITENSVENDA
   Dim TbPgto     As Object 'TB_CPGTOSVENDA
   Dim TbFLAN     As Object 'TB_FLAN
                                 
   '* Cliente
   Call SalvarCliente
   
   '* Atendimento
   Call PopulaClasse_Venda(TbVenda)
   bResult = TbVenda.Salvar
   
   If bResult Then
      mvarIDVENDA = TbVenda.IDVENDA
      
      Set cQueries = New Collection
      
      '*Atendimentos
      If GetTag(mvarMe.GrdAtendimento, "ISDIRT", 0) = 1 Then
            
         Set TBAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_VENDA")
         Set TBAtend.xDb = mvarSys.xDb
         TBAtend.IDLOJA = mvarIDLOJA
         TBAtend.IDVENDA = mvarIDVENDA
   
         cQueries.Add TBAtend.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA)
         With mvarMe.GrdAtendimento
            For i = 1 To .RowCount - 1
               If Val(.CellValue(i, "IDATENDIMENTO")) <> 0 Then
                  TBAtend.IDLOJA = mvarIDLOJA
                  TBAtend.IDVENDA = mvarIDVENDA
                  TBAtend.IDATENDIMENTO = .CellText(i, "IDATENDIMENTO")

                  cQueries.Add TBAtend.QrySave
               End If
            Next
         End With
         cQueries.Add TBAtend.QrySave
      End If
      
      '* Itens
      If GetTag(mvarMe.GrdVenda, "ISDIRT", 0) = 1 Then
         Set TbItens = CriarObjeto("BANCO_3R.TB_CITENSVENDA")
         Set TbItens.xDb = mvarSys.xDb
         TbItens.IDLOJA = mvarIDLOJA
         TbItens.IDVENDA = mvarIDVENDA
   
         cQueries.Add TbItens.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA)
         With mvarMe.GrdVenda
            For i = 1 To .RowCount - 1
               If xVal(.CellValue(i, "Total")) <> 0 Then
                  TbItens.IDLOJA = mvarIDLOJA
                  TbItens.IDVENDA = mvarIDVENDA
                  
                  TbItens.IDPROD = Val(.CellValue(i, "IDPROD"))
                  TbItens.NMPROD = .CellText(i, "IDPROD")
                  TbItens.QTDVENDA = xVal(.CellValue(i, "QTDVENDA"))
                  TbItens.UNIDCONTROLE = .CellValue(i, "UNIDCONTROLE")
                  TbItens.VLUNIT = xVal(.CellValue(i, "VLUNIT"))
                                    
                  cQueries.Add TbItens.QrySave
               End If
            Next
         End With
      End If
      
      '* Pagamentos
      If GetTag(mvarMe.GrdPagamento, "ISDIRT", 0) = 1 Then
         
         With mvarMe.GrdPagamento
            
            For i = 1 To .RowCount - 1
            
               Set TbPgto = CriarObjeto("BANCO_3R.TB_CPGTOSVENDA")
               Set TbPgto.xDb = mvarSys.xDb
               bAchou = TbPgto.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA, Ch_IDPGTO:=.CellValue(i, "IDPGTO"))
            
               If xVal(.CellValue(i, "VLPGTO")) <> 0 Then
                  sIn = IIf(Trim(sIn) = "", "", ", ") + sIn
                  
                  TbPgto.IDLOJA = mvarIDLOJA
                  TbPgto.IDVENDA = mvarIDVENDA
                  
                  TbPgto.IDFORMAPGTO = Val(.CellValue(i, "IDFORMAPGTO"))
                  TbPgto.IDCONDPGTO = Val(.CellValue(i, "IDCONDPGTO"))
                  TbPgto.VLPGTO = xVal(.CellValue(i, "VLPGTO"))
                  TbPgto.DTPGTO = mvarMe.CmbDTVENDA.Value
                  If bAchou Then
                     cQueries.Add TbPgto.QryUpdate
                     
                     tParcela = IIf(TbPgto.IDCONDPGTO > 1, TbPgto.IDCONDPGTO - 1, 1)
                     For iParcela = 1 To tParcela
                        Sql = " IDLOJA = " & SqlNum(mvarIDLOJA)
                        Sql = Sql & " And IDVENDA = " & SqlNum(mvarIDVENDA)
                        Sql = Sql & " And IDPGTO = " & SqlNum(.CellValue(i, "IDPGTO"))
                        Sql = Sql & " And NPARCELA = " & SqlNum(iParcela)
                        
                        Set TbFLAN = CriarObjeto("BANCO_3R.TB_FLAN")
                        With TbFLAN
                           Set .xDb = mvarSys.xDb
                           If .Pesquisar(Ch_Where:=Sql) Then
                              If ExibirPergunta("Existe lançamento associado ao pagamento." & vbNewLine & vbNewLine & "Deseja atualizar lançamento financeiro?") Then
                                 Dim TbConta As Object 'TB_FCCORRENTE
                                 If DefineContaLAN(TbConta, TbPgto.IDFORMAPGTO, mvarSys) Then
                                    If .AlterarChave(TbFLAN.IDLOJA, TbConta.IDCONTA) Then
                                       TbFLAN.IDCONTA = TbConta.IDCONTA
                                    End If
                                 End If
                                 .Valor = TbPgto.VLPGTO - IIf(TbPgto.IDFORMAPGTO = 1, TbVenda.VLTROCO, 0)
                                 Call DefineParcela(TbFLAN, TbVenda.DTVENDA, TbPgto.IDFORMAPGTO, TbPgto.IDCONDPGTO, TbPgto.VLPGTO, iParcela)
                                 cQueries.Add TbFLAN.QrySave
                              End If
                           End If
                        End With
                     Next
                  Else
                     cQueries.Add TbPgto.QryInsert
                  End If
               End If
            Next
            If Trim(sIn) <> "" Or .RowCount <= 1 Then
               If Trim(sIn) = "" Then sIn = "0"
               
               Sql = "Update FLAN"
               Sql = Sql & " Set IDVENDA=Null"
               Sql = Sql & " , IDPGTO =Null"
               Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
               Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
               Sql = Sql & " And Not IDPGTO In (" & sIn & ")"
               cQueries.Add Sql
               
               Sql = "Delete From CPGTOSVENDA "
               Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
               Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
               Sql = Sql & " And Not IDPGTO In (" & sIn & ")"
               cQueries.Add Sql
            End If
         End With
      End If
            
      SalvarVenda = bResult
      If cQueries.Count > 0 Then
         If mvarSys.GetParam(pCODPARAM:="NFE", Default:="0", pDescricao:="Nota Fiscal Eletrônica", pSECAO:="NFE") = 1 Then  'pCODSIS:="GLOBAL"
            Call SalvarRPS(cQueries)
         End If
         If mvarSys.xDb.Executa(cQueries) Then
            SalvarVenda = True
            Call SalvarEvento
         Else
            Call ExibirInformacao("Erro ao gravar itens e pagamentos.")
         End If
      End If
   Else
      Call ExibirInformacao("Erro ao gravar venda.")
   End If
        
Saida:
   Set TbVenda = Nothing
   Set TbItens = Nothing
   Set TbPgto = Nothing
   Set cQueries = Nothing
End Function
Private Sub SalvarRPS(Optional pQueries As Collection)
   Dim TBRecibo As Object 'TB_FRECIBO
   Dim bExiste As Boolean
   Dim Sql As String
         
   Set TBRecibo = CriarObjeto("BANCO_3R.TB_FRECIBO")
   With TBRecibo
      Set .xDb = mvarSys.xDb
      Sql = " IDVENDA=" & SqlNum(mvarIDVENDA)
      Sql = Sql & " And SERIE=1"
      bExiste = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=Sql)
      If xVal(mvarMe.TxtVLVENDA.Text) - xVal(mvarMe.TxtVLDESC.Text) <= 0 Then
         If bExiste Then
            .Status = 2
            pQueries.Add .QryUpdate
         Else
            Exit Sub
         End If
      End If
      
      .IDLOJA = mvarIDLOJA
      .IDVENDA = mvarIDVENDA
      .DTEMISSAO = mvarMe.CmbDTVENDA.Value
      .VLTOTAL = xVal(mvarMe.TxtVLVENDA.Text) - xVal(mvarMe.TxtVLDESC.Text)
      .Tipo = 0            '0-RPS 1-RPS-M (Conjugada/Mista)
      .SERIE = 1           '1-Serviço 2-Produto 3-Remessa
      .Status = 1          '1-Normal 2-Cancelado
      .NATUREZAOP = 0
      .ASSINATURA = ""
      .REGESPECIALTRIB = 0 '0– Nenhum; 1–Microempresa Municipal; 2–Estimativa; 3–Sociedade de Profissionais; 4–Cooperativa; 5–Microempreendedor Individual (MEI); 6–Art. 33, inc. II, item 8, Lei nº 691/84.
      .SIMPLES = 1         '0-Não-Optante pelo Simples Nacional; 1-Optante pelo Simples Nacional
      .TRIBUTACAO = 1      '1-Tributação no Município; 2-Tributação Fora do Município; 3-Operação Isenta; 4-Operação Imune; 5-Operação Suspensa por Decisão Judicial; 6-Operação Suspensa por Decisão Administrativa
      .INCENTIVOCULT = 0   '0–Não; 1–Sim.
      .CODSERVFEDERAL = "9602"               '4 dígitos
      .CODSERVMUNIC = "00000000000000072010" '20 dígitos
      
      '.IDRECIBO
      '.IDLOTE
      '.IDRECIBO0
      '.SERIE0
      
      'If .IDCLIENTE = 0 Then .DTCADASTRO = mvarSys.xDb.sysdate
      If .isDirt Then
         pQueries.Add .QrySave
      End If
   End With
   Set TBRecibo = Nothing
End Sub
Private Sub SalvarCliente()
   Dim TBCliente As Object 'TB_OCLIENTE
   Dim bClienteOK As Boolean
   
   Set TBCliente = CriarObjeto("BANCO_3R.TB_OCLIENTE")
   With TBCliente
      Set .xDb = mvarSys.xDb
      Call .Pesquisar(Ch_IDCLIENTE:=mvarIDCLIENTE, Ch_IDLOJA:=mvarIDLOJA)
      .IDLOJA = mvarIDLOJA
      .NOME = Trim(mvarMe.TxtNOME.Text)
      If .TEL2 <> Trim(mvarMe.TxtTEL1.Text) Then
         .TEL1 = Trim(mvarMe.TxtTEL1.Text)
      End If
      If .IDCLIENTE = 0 Then .DTCADASTRO = mvarSys.xDb.sysdate
      bClienteOK = (.NOME <> "") And (.TEL1 & .TEL2 <> "")
      If .isDirt And bClienteOK Then
         Call .Salvar
         mvarIDCLIENTE = .IDCLIENTE
      End If
   End With
   Set TBCliente = Nothing
   
End Sub
Private Sub SalvarEvento(Optional bDelete As Boolean)
   Dim pEvent     As CalendarEvent
   Dim oRs        As Object
   Dim Sql        As String
   Dim nIDATEND   As Long
   Dim sSITATEND  As String
   Dim nIDVENDA   As Long
   Dim sSITVENDA  As String
   Dim nIDCLIENTE As Long
   
   
   
   If mvarCalControl Is Nothing Then Exit Sub
   
   Sql = "Select V.IDVENDA, V.SITVENDA, A.IDATENDIMENTO, A.SITATEND, A.IDEVENTO, V.IDCLIENTE"
   Sql = Sql & " From CVENDA V "
   Sql = Sql & " Join OATENDIMENTO_VENDA AV On V.IDLOJA=AV.IDLOJA And V.IDVENDA=AV.IDVENDA"
   Sql = Sql & " Join OATENDIMENTO A On A.IDLOJA=AV.IDLOJA And A.IDATENDIMENTO=AV.IDATENDIMENTO"
   Sql = Sql & " Where V.IDLOJA=" & SqlNum(mvarIDLOJA)
   Sql = Sql & " AND V.IDVENDA=" & SqlNum(mvarIDVENDA)
   If mvarSys.xDb.AbreTabela(Sql, oRs) Then
      While Not oRs.EOF
         If mvarCalEvent Is Nothing Then
            Set pEvent = mvarCalControl.DataProvider.GetEvent(oRs("IDEVENTO") & "")
         Else
            Set pEvent = mvarCalEvent
         End If
         If Not pEvent Is Nothing Then
            nIDATEND = xVal(oRs("IDATENDIMENTO") & "")
            sSITATEND = Trim(oRs("SITATEND") & "")
            nIDVENDA = xVal(oRs("IDVENDA") & "")
            sSITVENDA = Trim(oRs("SITVENDA") & "")
            nIDCLIENTE = xVal(oRs("IDCLIENTE") & "")
            If bDelete Then
               nIDVENDA = 0
               sSITVENDA = ""
            End If
         
            Call SetCustomEvent(pEvent, mvarSys, pIDATENDIMENTO:=nIDATEND, _
                                                 pSITATEND:=sSITATEND, _
                                                 pIDVENDA:=nIDVENDA, _
                                                 pSITVENDA:=sSITVENDA, _
                                                 pIDCLIENTE:=nIDCLIENTE)
            
            Call mvarCalControl.DataProvider.ChangeEvent(pEvent)
            Call RefreshEvent(mvarSys, mvarCalControl, pEvent, False)

         End If
         oRs.MoveNext
      Wend
   End If
End Sub
Private Sub PopulaClasse_Venda(ByRef pTBVenda As Object)
   Dim sDate   As String
   
   Set pTBVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
   
   With pTBVenda
      .xDb = mvarSys.xDb
      
      Call .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA)
      .IDVENDA = mvarIDVENDA
      .IDLOJA = mvarIDLOJA
      
      .DTVENDA = mvarMe.CmbDTVENDA.Value
      .IDCLIENTE = mvarIDCLIENTE
      '.IDFUNCIONARIO = mvaridfuncionario
      .IDMOEDA = IIf(Val(Sys.Propriedades("IDMOEDA")) = 0, 1, Sys.Propriedades("IDMOEDA"))
      If .SITVENDA = "" Then
         .SITVENDA = "00"
      End If
      .VLDESC = xVal(mvarMe.TxtVLDESC.Text)
      .VLPGTO = xVal(mvarMe.TxtVLPGTO.Text)
      .VLTROCO = ValorReal(mvarMe.TxtVLTROCO.Text)
      .VLVENDA = xVal(mvarMe.TxtVLVENDA.Text)
      '.OBS = mvarMe.TxtOBS.Text
   End With
End Sub
Private Sub mvarMe_GrdAtendimentoColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   bDoDefault = False
End Sub
Private Sub mvarMe_GrdAtendimentoLostFocus()
   Dim sTag As String
   Dim i    As Integer
   Dim j    As Integer

   With mvarMe.GrdAtendimento
      '.Tag = SetTag(.Tag, "ISDIRT", 0)
      If .RowCount <= 1 Then .Tag = SetTag(.Tag, "ISDIRT", 1)
      For i = 1 To .RowCount - 1
         If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
         For j = 1 To .ColCount
            If GetTag(sTag, .ColKey(j), 0) <> .CellValue(i, j) Then .Tag = SetTag(.Tag, "ISDIRT", 1)
            If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
         Next
         If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
      Next
   End With
End Sub

Private Sub mvarMe_GrdAtendimentoMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
   With mvarMe.GrdAtendimento
      If lRow = .RowCount Then
         If bVendaFechada Then
            Call ExibirInformacao("Venda fechada.")
            Exit Sub
         End If
      
         If lRow > 1 Then
            If Val(.CellValue(lRow - 1, "IDATENDIMENTO")) = 0 Then Exit Sub
         End If
         Call IncluiGrdLinhaAtendimento(True)
         
         If lRow > 0 Then
            If Val(.CellValue(lRow, "IDATENDIMENTO")) = 0 Then
               Call .SetCurCell(lRow, "IDATENDIMENTO")
            End If
         End If
      End If
   End With

End Sub

Private Sub mvarMe_GrdAtendimentoRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   If bVendaFechada Then
      bCancel = True
      Exit Sub
   End If
   
   With mvarMe.GrdAtendimento
      bCancel = (GetTag(.ColTag(lCol), "EDIT", True) = False)
      If .CellFmtString(lRow, lCol) <> "" Then sText = .CellValue(lRow, lCol)
      eTextEditOpt = GetTag(.ColTag(lCol), "EDITOPT", 0)
   End With
End Sub
Private Sub mvarMe_GrdAtendimentoValidate(Cancel As Boolean)
   mvarMe.GrdAtendimento.CommitEdit
End Sub
Private Sub mvarMe_GrdPagamentoAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   Dim Sql As String
   Dim TbCond  As Object 'TB_CCONDPGTO
   Dim sAux    As String
   
   Set TbCond = CriarObjeto("BANCO_3R.TB_CCONDPGTO")
   Set TbCond.xDb = mvarSys.xDb
   
   With mvarMe.GrdPagamento
      If .ColKey(lCol) = "IDFORMAPGTO" Then
         Sql = "ATIVO=1"
'         Select Case Val(.CellValue(lRow, "IDFORMAPGTO"))
'            Case 1: Sql = Sql & " And IDCONDPGTO=1"
'            Case 2: Sql = Sql & " And IDCONDPGTO=1"
'            Case 3: Sql = Sql & " And IDCONDPGTO>1"
'            Case 4: Sql = Sql
'         End Select
         
         With .Combos("IDCONDPGTO")
            If TbCond.Pesquisar(Ch_Where:=Sql, Ch_ORDERBY:="IDCONDPGTO") Then
               .Clear
               While Not TbCond.Rs.EOF
                  TbCond.Popula
                  sAux = TbCond.CONDPGTO
                  sAux = sAux & Space(15 - Len(sAux))
                  .AddItem sAux, CStr(TbCond.IDCONDPGTO)
   
                  TbCond.Rs.MoveNext
               Wend
            End If
            Set TbCond = Nothing
            
            If .ListCount > 0 Then
               If InArray(Val(mvarMe.GrdPagamento.CellValue(lRow, "IDFORMAPGTO")), Array(1, 2, 4)) Then
                  mvarMe.GrdPagamento.CellValue(lRow, "IDCONDPGTO") = .ItemValue(0)
               ElseIf Val(mvarMe.GrdPagamento.CellValue(lRow, "IDFORMAPGTO")) = 3 Then
                  mvarMe.GrdPagamento.CellValue(lRow, "IDCONDPGTO") = .ItemValue(1)
               End If
            End If
         End With
      ElseIf .ColKey(lCol) = "VLPGTO" Then
         .CellValue(lRow, "VLPGTO") = ValBr(.CellValue(lRow, "VLPGTO"))
         Call PopulaValores
      End If
   End With
End Sub
Private Sub PopulaValores()
   Dim i       As Integer
   Dim nVlPgto As Currency
   Dim nVlAPagar As Currency
   
   With mvarMe
      With .GrdPagamento
         nVlPgto = 0
         For i = 1 To .RowCount
            nVlPgto = nVlPgto + xVal(.CellValue(i, "VLPGTO"))
         Next
      End With
      nVlAPagar = xVal(.TxtVLVENDA.Text)
      nVlAPagar = nVlAPagar - xVal(.TxtVLDESC.Text)
      nVlAPagar = nVlAPagar - nVlPgto
      
      .TxtVLPGTO.Text = ValBr(nVlPgto)
      Call PopulaTroco(-1 * nVlAPagar)
   End With
End Sub
Private Sub PopulaTroco(nVlAPagar As Currency)
   With mvarMe
      If nVlAPagar < 0 Then
         .TxtVLTROCO.Text = "(" & ValBr(nVlAPagar) & ")"
         .LblTroco.Caption = "A PAGAR"
         .TxtVLTROCO.ForeColor = vbRed
      ElseIf nVlAPagar > 0 Then
         .TxtVLTROCO.Text = ValBr(nVlAPagar)
         .LblTroco.Caption = "TROCO"
         .TxtVLTROCO.ForeColor = vbBlue
      Else
         .TxtVLTROCO.Text = ValBr(0)
         .TxtVLTROCO.ForeColor = vbBlack
      End If
   End With
End Sub
Private Sub mvarMe_GrdPagamentoBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   If lConvErr <> 0 Then
      ExibirAviso "Valor inválido."
      eResult = igEditResProceed
   End If
   With mvarMe.GrdPagamento
      If .ColKey(lCol) = "IDCONDPGTO" Then
         If InArray(Val(.CellValue(lRow, "IDFORMAPGTO")), Array(1, 2, 4)) Then
            If vNewValue <> 1 Then
               Call ExibirAviso("Condição Inválida")
               eResult = igEditResCancel
            End If
         ElseIf Val(.CellValue(lRow, "IDFORMAPGTO")) = 3 Then
            If vNewValue = 1 Then
               Call ExibirAviso("Condição Inválida")
               eResult = igEditResCancel
            End If
            
         End If
      End If
   End With
End Sub
Private Sub mvarMe_GrdPagamentoColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   bDoDefault = False
End Sub
Private Sub mvarMe_GrdPagamentoLostFocus()
   Dim sTag As String
   Dim i    As Integer
   Dim j    As Integer

   With mvarMe.GrdPagamento
      '.Tag = SetTag(.Tag, "ISDIRT", 0)
      If .RowCount <= 1 Then .Tag = SetTag(.Tag, "ISDIRT", 1)
      For i = 1 To .RowCount - 1
         If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
         For j = 1 To .ColCount
            If GetTag(sTag, .ColKey(j), 0) <> .CellValue(i, j) Then .Tag = SetTag(.Tag, "ISDIRT", 1)
            If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
         Next
         If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
      Next
   End With
End Sub

Private Sub mvarMe_GrdPagamentoMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)

   With mvarMe.GrdPagamento
      If lRow = .RowCount Then
         If bVendaFechada Then
            Call ExibirInformacao("Venda fechada.")
            Exit Sub
         End If
         
         If .RowCount > .Combos("IDFORMAPGTO").ListCount Then
            Call ExibirAviso("Não existem forma de pagamento configurada.")
            Exit Sub
         End If
   
         If lRow > 1 Then
            If xVal(.CellValue(lRow - 1, "VLPGTO")) = 0 Then Exit Sub
         End If
         Call IncluiGrdLinhaPgto
   
         If xVal(.CellValue(lRow, "VLPGTO")) = 0 Then
            Call .SetCurCell(lRow, "VLPGTO")
         End If
      End If
   End With
End Sub
Private Sub mvarMe_GrdPagamentoRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   If bVendaFechada Then
      bCancel = True
      Exit Sub
   End If
   
   With mvarMe.GrdPagamento
      bCancel = (GetTag(.ColTag(lCol), "EDIT", True) = False)
      If .CellFmtString(lRow, lCol) <> "" Then sText = .CellValue(lRow, lCol)
      eTextEditOpt = GetTag(.ColTag(lCol), "EDITOPT", 0)
   End With
End Sub
Private Sub mvarMe_GrdPagamentoValidate(Cancel As Boolean)
   mvarMe.GrdPagamento.CommitEdit
End Sub
Private Sub mvarMe_GrdVendaAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   Dim MyPROD  As Object 'TB_SPRODUTO
   
   With mvarMe.GrdVenda
      If .ColKey(lCol) = "IDPROD" Then
         Set MyPROD = CriarObjeto("BANCO_3R.TB_SPRODUTO")
         Set MyPROD.xDb = mvarSys.xDb
         If MyPROD.Pesquisar(Ch_IDPROD:=.CellValue(lRow, "IDPROD")) Then
            .CellValue(lRow, "UNIDCONTROLE") = MyPROD.UNIDCONTROLE
            .CellValue(lRow, "VLUNIT") = ValBr(MyPROD.VLVENDA)
            .CellValue(lRow, "Total") = ValBr(xVal(.CellValue(lRow, "QTDVENDA")) * xVal(.CellValue(lRow, "VLUNIT")))
            Call CalculaVenda
            Call PopulaValores
         End If
         Set MyPROD = Nothing
         
      ElseIf .ColKey(lCol) = "QTDVENDA" Then
         .CellValue(lRow, "Total") = ValBr(xVal(.CellValue(lRow, "QTDVENDA")) * xVal(.CellValue(lRow, "VLUNIT")))
         Call CalculaVenda
         Call PopulaValores
      End If
   End With
End Sub
Private Sub mvarMe_GrdVendaBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   If lConvErr <> 0 Then
      ExibirAviso "Valor inválido."
      eResult = igEditResProceed
   End If
End Sub
Private Sub mvarMe_GrdVendaColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   bDoDefault = False
End Sub
Private Sub mvarMe_GrdVendaLostFocus()
   Dim sTag As String
   Dim i    As Integer
   Dim j    As Integer
   
   With mvarMe.GrdVenda
      '.Tag = SetTag(.Tag, "ISDIRT", 0)
      If .RowCount <= 1 Then .Tag = SetTag(.Tag, "ISDIRT", 1)
      For i = 1 To .RowCount - 1
         If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
         For j = 1 To .ColCount
            If GetTag(sTag, .ColKey(j), 0) <> .CellValue(i, j) Then .Tag = SetTag(.Tag, "ISDIRT", 1)
            If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
         Next
         If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
      Next
   End With
End Sub

Private Sub mvarMe_GrdVendaMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)

   With mvarMe.GrdVenda
      If lRow = .RowCount Then
         If bVendaFechada Then
            Call ExibirInformacao("Venda fechada.")
            Exit Sub
         End If
         
         If .RowCount > .Combos("IDPROD").ListCount Then
            Call ExibirAviso("Não existem mais produtos de venda. (" + CStr(.Combos("IDPROD").ListCount) + ")")
            Exit Sub
         End If
         
         If lRow > 1 Then
            If Val(.CellValue(lRow - 1, "IDPROD")) = 0 Then Exit Sub
            If xVal(.CellValue(lRow - 1, "QTDVENDA")) = 0 Then Exit Sub
         End If
         Call IncluiGrdLinhaVenda
         
         If Val(.CellValue(lRow, "IDPROD")) = 0 Then
            Call .SetCurCell(lRow, "IDPROD")
         Else
            Call .SetCurCell(lRow, "QTDVENDA")
         End If
      End If
   End With
End Sub
Private Sub mvarMe_GrdVendaRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   If bVendaFechada Then
      bCancel = True
      Exit Sub
   End If
   With mvarMe.GrdVenda
      bCancel = (GetTag(.ColTag(lCol), "EDIT", True) = False)
      If .CellFmtString(lRow, lCol) <> "" Then sText = .CellValue(lRow, lCol)
      eTextEditOpt = GetTag(.ColTag(lCol), "EDITOPT", 0)
   End With
End Sub

Private Sub mvarMe_GrdVendaValidate(Cancel As Boolean)
   mvarMe.GrdVenda.CommitEdit
End Sub

Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim lRow As Long
   
   If KeyCode = vbKeyDelete Then
      With mvarMe
         If .ActiveControl Is .GrdPagamento _
         Or .ActiveControl Is .GrdVenda _
         Or (.ActiveControl Is .GrdAtendimento And .GrdAtendimento.CurRow > 1) Then
            With .ActiveControl
               lRow = .CurRow
               If lRow <> .RowCount And lRow > 0 Then
                  .RemoveRow (.CurRow)
                  .Tag = SetTag(.Tag, "ISDIRT", 1)
                  .SetFocus
                  If lRow < .RowCount Then
                     .SetCurCell lRow, 1
                  ElseIf lRow > 1 Then
                     .SetCurCell lRow - 1, 1
                  End If
               End If
               
               If mvarMe.ActiveControl Is mvarMe.GrdAtendimento Then
                  Call PopulaGRDPRODUTO
               Else
                  For lRow = 1 To .RowCount - 1
                     .CellValue(lRow, 1) = lRow
                  Next
               End If
            End With
            Call CalculaVenda
            Call PopulaValores
         End If
      End With
   End If
End Sub
Private Sub mvarMe_Load()
   mvarMe.Icon = mvarSys.MDI.Icon
   '* Monta Tela
   Call MontaTela
   Call MontaGrdVenda
   Call MontaGrdPgto
   Call MontaGrdAtendimento
   
   Call CarregaFechamento
   Call PopulaTela
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub CarregaFechamento()
   Dim TbVenda As Object 'TB_CVENDA
   Dim TbAV    As Object 'TB_OATENDIMENTO_VENDA
   
   bVendaFechada = False
   
   If mvarIDVENDA = 0 And mvarIDATENDIMENTO <> 0 Then
      Set TbAV = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_VENDA")
      Set TbAV.xDb = mvarSys.xDb
      
      If TbAV.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO) Then
         mvarIDVENDA = TbAV.IDVENDA
      End If
   End If
   If mvarIDVENDA <> 0 Then
      Set TbVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
      Set TbVenda.xDb = mvarSys.xDb
      If TbVenda.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA) Then
         bVendaFechada = (TbVenda.SITVENDA = gSITVENDA_Fechada)
      End If
   End If
   Set TbVenda = Nothing
   Set TbAV = Nothing
End Sub
Private Sub MontaTela()
   With mvarMe
      .CmdExcluir.Enabled = False
      .TxtTEL1.MaxLength = mvarSys.Propriedades("MaxLenTel")
      
      .TxtIDVENDA.BackColor = &HCCEEDD
      .CmbDTVENDA.Enabled = False
   
      .TxtVLVENDA.Font.Bold = True
      .TxtVLVENDA.Font.Size = 16
      .TxtVLVENDA.Font.Name = "Courier New"
      .TxtVLVENDA.BackColor = .TxtIDVENDA.BackColor
      .TxtVLVENDA.ForeColor = vbBlack
      .TxtVLVENDA.Enabled = True
      .TxtVLVENDA.Locked = True
      
      .TxtVLDESC.Font.Size = 16
      .TxtVLDESC.Font.Name = "Courier New"
      .TxtVLDESC.Locked = False
      
      .TxtVLPGTO.Font.Size = 16
      .TxtVLPGTO.Font.Name = "Courier New"
      .TxtVLPGTO.Enabled = True
      .TxtVLPGTO.Locked = True
      
      .TxtVLTROCO.Font.Size = 16
      .TxtVLTROCO.Font.Name = "Courier New"
      .TxtVLTROCO.Locked = True
            
   End With
End Sub
Private Sub PopulaTela()
   Dim Sql     As String
   Dim TBAtend As Object 'TB_OATENDIMENTO
   Dim TbVenda As Object 'TB_CVENDA
   Dim TbAV    As Object 'TB_OATENDIMENTO_VENDA
   Dim bOk     As Boolean
   
   Call LimparTela
   
   If mvarIDVENDA = 0 And mvarIDATENDIMENTO <> 0 Then
      Set TbAV = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_VENDA")
      Set TbAV.xDb = mvarSys.xDb
      
      bOk = TbAV.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO)
      If bOk Then mvarIDVENDA = TbAV.IDVENDA
   End If
   
   If mvarIDVENDA > 0 Then
      If TbVenda Is Nothing Then
         Set TbVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
         Set TbVenda.xDb = mvarSys.xDb
         bOk = TbVenda.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA)
      End If
      If bOk Then
         With TbVenda
            'bVendaFechada = (.SITVENDA = gSITVENDA_Fechada)
            mvarMe.CmdFechar.Enabled = Not bVendaFechada
            mvarMe.CmdOk.Enabled = Not bVendaFechada
            mvarMe.CmdExcluir.Enabled = Not bVendaFechada
            'mvarMe.CmdIDCLIENTE.Enabled = Not bVendaFechada
            mvarMe.CmdLovCli.Enabled = Not bVendaFechada
            mvarMe.TxtNOME.Locked = bVendaFechada
            mvarMe.TxtTEL1.Locked = bVendaFechada
            mvarMe.CmdChave.Visible = bVendaFechada
                     
            mvarMe.TxtIDVENDA.Text = StrZero(.IDVENDA, mvarSys.Propriedades("LENIDVENDA"))
            mvarMe.CmbDTVENDA.Value = .DTVENDA
                                    
            Call PopulaCliente(.IDCLIENTE)

            mvarMe.TxtVLVENDA.Text = ValBr(.VLVENDA)
            mvarMe.TxtVLDESC.Text = ValBr(.VLDESC)
            mvarMe.TxtVLPGTO.Text = ValBr(.VLPGTO)
            Call PopulaTroco(.VLTROCO)
            
'            Call LocalizarCombo(mvarMe.CmbIDFUNCIONARIO, .IDFUNCIONARIO, True, True)
'            mvarMe.TxtOBS.Text = .OBS

            Call PopulaGRDAtendimento
            Call PopulaGRDVENDA
            Call PopulaGRDPGTO
            Call PopulaRecibo
         End With
      End If
      Set TbVenda = Nothing
   ElseIf mvarIDATENDIMENTO <> 0 Then
      Set TBAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
      Set TBAtend.xDb = mvarSys.xDb
      If TBAtend.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO) Then
         With TBAtend
            Call PopulaCliente(.IDCLIENTE)
            Call PopulaGRDAtendimento(mvarIDATENDIMENTO)
         End With
      End If
      Set TBAtend = Nothing
   End If
   Call CalculaVenda
   Call PopulaValores
End Sub
Private Sub PopulaRecibo()
   Dim MyTb As Object 'TB_FRECIBO
   Dim Sql  As String

   DoEvents
   Set MyTb = CriarObjeto("BANCO_3R.TB_FRECIBO")
   Set MyTb.xDb = mvarSys.xDb
   
   Sql = " IDVENDA=" & SqlNum(mvarIDVENDA)
   Sql = Sql & " And SERIE=1"
   If MyTb.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=Sql) Then
      mvarMe.CmdRecibo.Caption = "Recibo " & StrZero(MyTb.IDRECIBO, 6)
      mvarMe.CmdRecibo.Enabled = True
   Else
      mvarMe.CmdRecibo.Enabled = False
   End If
End Sub
Private Sub PopulaGRDVENDA()
   Dim MyTb As Object 'TB_CITENSVENDA
   Dim sTag As String
   Dim i    As Integer
   Dim Sql  As String

   DoEvents
   Set MyTb = CriarObjeto("BANCO_3R.TB_CITENSVENDA")
   Set MyTb.xDb = mvarSys.xDb
      
   
   If MyTb.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA, Ch_ORDERBY:="IDITEM") Then
      With mvarMe.GrdVenda
         .Clear
         Call IncluiGrdLinhaVenda
         
         While Not MyTb.Rs.EOF
            MyTb.Popula
                  
            Call IncluiGrdLinhaVenda
         
            i = .RowCount - 1
            If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
            
            Call SetTag(sTag, "IDITEM", MyTb.IDITEM)
            Call SetTag(sTag, "IDPROD", MyTb.IDPROD)
            Call SetTag(sTag, "QTDVENDA", MyTb.QTDVENDA)
            Call SetTag(sTag, "UNIDCONTROLE", MyTb.UNIDCONTROLE)
            Call SetTag(sTag, "VLUNIT", ValBr(MyTb.VLUNIT))
            Call SetTag(sTag, "Total", ValBr(MyTb.QTDVENDA * MyTb.VLUNIT))
            
            .RowTag(i) = sTag
            
            .CellValue(i, "IDITEM") = MyTb.IDITEM
            .CellValue(i, "IDPROD") = MyTb.IDPROD
            .CellValue(i, "QTDVENDA") = MyTb.QTDVENDA
            .CellValue(i, "UNIDCONTROLE") = MyTb.UNIDCONTROLE
            .CellValue(i, "VLUNIT") = ValBr(MyTb.VLUNIT)
            .CellValue(i, "Total") = ValBr(MyTb.QTDVENDA * MyTb.VLUNIT)
                     
            MyTb.Rs.MoveNext
         Wend
         .RowVisible(1) = True
         '.CellSelected(1, "FOTOTIPO") = True
         .Tag = SetTag(.Tag, "ISDIRT", 0)
      End With
   End If
   Set MyTb = Nothing
End Sub
Private Sub PopulaGRDPGTO()
   Dim MyTb    As Object 'TB_CPGTOSVENDA
   Dim MyTbP   As Object 'TB_SPRODUTO
   Dim sTag    As String
   Dim i       As Integer
   Dim Sql     As String

   DoEvents
   Set MyTbP = CriarObjeto("BANCO_3R.TB_SPRODUTO")
   Set MyTbP.xDb = mvarSys.xDb
   
   Set MyTb = CriarObjeto("BANCO_3R.TB_CPGTOSVENDA")
   Set MyTb.xDb = mvarSys.xDb
      
   
   If MyTb.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA, Ch_ORDERBY:="IDPGTO") Then
      With mvarMe.GrdPagamento
         .Clear
         Call IncluiGrdLinhaPgto
         
         While Not MyTb.Rs.EOF
            MyTb.Popula
            
            Call IncluiGrdLinhaPgto
         
            i = .RowCount - 1
            If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
            Call SetTag(sTag, "IDPGTO", MyTb.IDPGTO)
            Call SetTag(sTag, "IDCONDPGTO", MyTb.IDCONDPGTO)
            Call SetTag(sTag, "IDFORMAPGTO", MyTb.IDFORMAPGTO)
            Call SetTag(sTag, "VLPGTO", ValBr(MyTb.VLPGTO))
                        
            .RowTag(i) = sTag
            
            .CellValue(i, "IDPGTO") = MyTb.IDPGTO
            .CellValue(i, "IDFORMAPGTO") = CStr(MyTb.IDFORMAPGTO)
            Call mvarMe_GrdPagamentoAfterCommitEdit(i, .ColIndex("IDFORMAPGTO"))
            .CellValue(i, "IDCONDPGTO") = CStr(MyTb.IDCONDPGTO)
            .CellValue(i, "VLPGTO") = ValBr(MyTb.VLPGTO)
                     
            MyTb.Rs.MoveNext
         Wend
         .RowVisible(1) = True
         '.CellSelected(1, "FOTOTIPO") = True
         .Tag = SetTag(.Tag, "ISDIRT", 0)
      End With
   End If
   Set MyTb = Nothing
End Sub
Private Sub PopulaGRDPRODUTO()
   Dim MyTbServ As Object 'TB_OSESSAO
   Dim MyTbProd As Object 'TB_OATENDIMENTO_PRODUTO
   Dim sTag    As String
   Dim i       As Integer
   Dim Sql     As String
   
   Dim pIDATENDIMENTO As Long
   Dim MyItem  As Object 'TB_CITENSVENDA
   Dim MyPROD  As Object 'TB_SPRODUTO
   Dim nIDPROD As Long
   Dim CItens  As Collection
   Dim j       As Integer

   Set CItens = New Collection
   
   With mvarMe.GrdVenda
      For j = 1 To mvarMe.GrdAtendimento.RowCount
         pIDATENDIMENTO = Val(mvarMe.GrdAtendimento.CellValue(j, "IDATENDIMENTO"))
         
         Set MyTbServ = CriarObjeto("BANCO_3R.TB_OSESSAO")
         Set MyTbServ.xDb = mvarSys.xDb
         If MyTbServ.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=pIDATENDIMENTO) Then
            While Not MyTbServ.Rs.EOF
               MyTbServ.Popula
                                                      
               nIDPROD = GetProdutoVenda(MyTbServ.IDTPSERVICO, MyTbServ.IDTPTRATAMENTO, MyTbServ.IDAREA)
               If nIDPROD > 0 Then
                  If Not ExisteItem(CItens, CStr(nIDPROD)) Then
                     Set MyItem = CriarObjeto("BANCO_3R.TB_CITENSVENDA")
                     CItens.Add MyItem, CStr(nIDPROD)
                  End If
                  CItens(CStr(nIDPROD)).IDPROD = nIDPROD
                  CItens(CStr(nIDPROD)).QTDVENDA = CItens(CStr(nIDPROD)).QTDVENDA + 1
               End If
               MyTbServ.Rs.MoveNext
            Wend
         End If
         Set MyTbServ = Nothing
         
         Set MyTbProd = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_PRODUTO")
         Set MyTbProd.xDb = mvarSys.xDb
         If MyTbProd.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=pIDATENDIMENTO) Then
            While Not MyTbProd.Rs.EOF
               MyTbProd.Popula
                                                      
               nIDPROD = MyTbProd.IDPROD
               If nIDPROD > 0 Then
                  If Not ExisteItem(CItens, CStr(nIDPROD)) Then
                     Set MyItem = CriarObjeto("BANCO_3R.TB_CITENSVENDA")
                     CItens.Add MyItem, CStr(nIDPROD)
                  End If
                  CItens(CStr(nIDPROD)).IDPROD = nIDPROD
                  CItens(CStr(nIDPROD)).QTDVENDA = CItens(CStr(nIDPROD)).QTDVENDA + MyTbProd.QTDCOMPRA
               End If
               MyTbProd.Rs.MoveNext
            Wend
         End If
         Set MyTbProd = Nothing
      Next
      
      If CItens.Count > 0 Then
         .Clear
         Call IncluiGrdLinhaVenda
         
         For Each MyItem In CItens
            
            Call IncluiGrdLinhaVenda
            i = .RowCount - 1
            If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
            
            
            Set MyPROD = CriarObjeto("BANCO_3R.TB_SPRODUTO")
            Set MyPROD.xDb = mvarSys.xDb
            Call MyPROD.Pesquisar(MyItem.IDPROD)
         
            Call SetTag(sTag, "IDITEM", i)
            Call SetTag(sTag, "IDPROD", MyItem.IDPROD)
            Call SetTag(sTag, "QTDVENDA", MyItem.QTDVENDA)
            Call SetTag(sTag, "UNIDCONTROLE", MyPROD.UNIDCONTROLE)
            Call SetTag(sTag, "VLUNIT", ValBr(MyPROD.VLVENDA))
            Call SetTag(sTag, "Total", ValBr(xVal(.CellValue(i, "QTDVENDA")) * xVal(.CellValue(i, "VLUNIT"))))
            
            .RowTag(i) = sTag
            
            .CellValue(i, "IDITEM") = i
            .CellValue(i, "IDPROD") = MyItem.IDPROD
            .CellValue(i, "QTDVENDA") = MyItem.QTDVENDA
            .CellValue(i, "UNIDCONTROLE") = MyPROD.UNIDCONTROLE
            .CellValue(i, "VLUNIT") = ValBr(MyPROD.VLVENDA)
            .CellValue(i, "Total") = ValBr(xVal(.CellValue(i, "QTDVENDA")) * xVal(.CellValue(i, "VLUNIT")))
                     
         Next
         
         .RowVisible(1) = True
         '.CellSelected(1, "FOTOTIPO") = True
         .Tag = SetTag(.Tag, "ISDIRT", 0)
      End If
      If .RowCount > 1 Then
         Call SetTag(mvarMe.GrdVenda, "ISDIRT", 1)
      End If
   End With
End Sub
Private Sub PopulaGRDAtendimento(Optional pIDATENDIMENTO As Long = 0)
   Dim MyRs As Object
   Dim sTag As String
   Dim i    As Integer
   Dim Sql  As String

   DoEvents
   Sql = "Select Right(REPLICATE('0'," + CStr(mvarSys.Propriedades("LENIDVENDA")) + ")+ Cast(A.IDATENDIMENTO as varchar)," + CStr(mvarSys.Propriedades("LENIDVENDA")) + ") [IDATENDIMENTO]"
   Sql = Sql & " ,  C.TEL1, C.NOME"
   Sql = Sql & " From OATENDIMENTO A"
   Sql = Sql & " Join OCLIENTE C On A.IDLOJA=C.IDLOJA And A.IDCLIENTE=C.IDCLIENTE"
   If pIDATENDIMENTO = 0 Then
      Sql = Sql & " Join OATENDIMENTO_VENDA AV On A.IDLOJA=AV.IDLOJA And A.IDATENDIMENTO=AV.IDATENDIMENTO"
      Sql = Sql & " Where AV.IDVENDA = " & mvarIDVENDA
   Else
      Sql = Sql & " Where A.IDATENDIMENTO = " & pIDATENDIMENTO
   End If
   Sql = Sql & " Order By IDATENDIMENTO"
   
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      With mvarMe.GrdAtendimento
         .Clear
         Call IncluiGrdLinhaAtendimento(False)
         
         While Not MyRs.EOF
            
            Call IncluiGrdLinhaAtendimento(False)
            
            i = .RowCount - 1
            If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
            Call SetTag(sTag, "IDATENDIMENTO", StrZero(MyRs("IDATENDIMENTO") & "", mvarSys.Propriedades("LENIDVENDA")))
            Call SetTag(sTag, "TEL1", MyRs("TEL1") & "")
            Call SetTag(sTag, "IDCLIENTE", MyRs("NOME"))
                        
            .RowTag(i) = sTag
            
            .CellValue(i, "IDATENDIMENTO") = StrZero(MyRs("IDATENDIMENTO") & "", mvarSys.Propriedades("LENIDVENDA"))
            .CellValue(i, "TEL1") = MyRs("TEL1") & ""
            .CellValue(i, "IDCLIENTE") = MyRs("NOME") & ""
                     
            MyRs.MoveNext
         Wend
         .RowVisible(1) = True
         .Tag = SetTag(.Tag, "ISDIRT", 0)
      End With
      If pIDATENDIMENTO <> 0 Then
         Call PopulaGRDPRODUTO
         Call SetTag(mvarMe.GrdAtendimento, "ISDIRT", 1)
      End If
   End If
End Sub

Private Sub LimparTela()
   With mvarMe
      .TxtIDVENDA.Text = ""
      If mvarIDATENDIMENTO = 0 Then
         .TxtIDVENDA.Enabled = False
         .CmbDTVENDA.Value = mvarSys.xDb.sysdate()
      Else
         Dim TBAtend As Object 'TB_OATENDIMENTO
         
         Set TBAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
         With TBAtend
            Set .xDb = mvarSys.xDb
            mvarMe.TxtIDVENDA.Enabled = False
            If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO) Then
               mvarMe.CmbDTVENDA.Value = .DTATEND
               If Format(mvarMe.CmbDTVENDA.Value, "DD/MM/YYYY") <> dToday Then
                  mvarMe.TxtIDVENDA.Enabled = True
               End If
            Else
               mvarMe.CmbDTVENDA.Value = mvarSys.xDb.sysdate()
            End If
            Set TBAtend = Nothing
         End With
      End If
      
      Call SetTag(.TxtNOME, "IDCLIENTE", "")
      .TxtTEL1.Text = ""
      .TxtNOME.Text = ""
      
      .TxtVLDESC.Text = ValBr(0)
      .TxtVLPGTO.Text = ValBr(0)
      .TxtVLTROCO.Text = ValBr(0)
      .TxtVLVENDA.Text = ValBr(0)
      '.TxtOBS.Text = ""
      'If .CmbIDFUNCIONARIO.ListCount > 1 Then .CmbIDFUNCIONARIO.ListIndex = -1
            
      .CmdRecibo.Caption = "Recibo 000000"
      .CmdRecibo.Enabled = False
            
      .GrdPagamento.Clear False
      .GrdVenda.Clear False
      .GrdAtendimento.Clear False
                  
      Call IncluiGrdLinhaVenda
      Call IncluiGrdLinhaPgto
      Call IncluiGrdLinhaAtendimento
   End With
End Sub
Private Sub MontaGrdVenda()
   Dim TbProd  As Object 'TB_SPRODUTO
   Dim Sql     As String
   Dim sAux    As String
   Dim i       As Integer
   
   Set TbProd = CriarObjeto("BANCO_3R.TB_SPRODUTO")
   Set TbProd.xDb = mvarSys.xDb
   
   With mvarMe.GrdVenda
      .ShowControlsInAllCells = True
      .Header.DragCols = False
      .SilentValidation = True
      '.Header.ImageList = mvarMe.IlsIcons
      '.ImageList = mvarMe.IlsIcons
      With .Combos
         With .Add("IDPROD")
            Sql = "EVENDA=1"
            Sql = Sql & " And SITPROD='A'"
            If TbProd.Pesquisar(Ch_Where:=Sql, Ch_ORDERBY:="CODPROD") Then

               While Not TbProd.Rs.EOF
                  TbProd.Popula
                  sAux = TbProd.NMPROD
                  If Len(sAux) < 10 Then
                     sAux = sAux & Space(10 - Len(sAux))
                  End If
                  .AddItem sAux, TbProd.IDPROD
                  
                  TbProd.Rs.MoveNext
               Wend
            End If
            Set TbProd = Nothing
            .AutoAdjustWidth
         End With
      End With
      
      
      With .AddCol(sKey:="IDITEM", sHeader:="#", lWidth:=25, bVisible:=True)
         .eType = igCellText
         .sCtrlKey = "IDITEM"
      End With
      With .AddCol(sKey:="IDPROD", sHeader:="Produto", lWidth:=110, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDPROD"
      End With
      With .AddCol(sKey:="QTDVENDA", sHeader:="Qtd.", lWidth:=40, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextCenter
         .sCtrlKey = "QTDVENDA"
      End With
      With .AddCol(sKey:="UNIDCONTROLE", sHeader:="Unid.", lWidth:=40, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextCenter
         .sCtrlKey = "UNIDCONTROLE"
      End With
      With .AddCol(sKey:="VLUNIT", sHeader:="Valor Unit.", lWidth:=65, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight
         '.sFmtString = "R$ #,###.00"
         .sCtrlKey = "VLUNIT"
      End With
      With .AddCol(sKey:="Total", sHeader:="Total", lWidth:=65, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight
         .sCtrlKey = "Total"
      End With
            
      For i = 1 To .ColCount
         .ColTag(i) = ""
      Next
            
      .ColTag("IDITEM") = SetTag(.ColTag("IDITEM"), "EDIT", False)
      .ColTag("UNIDCONTROLE") = SetTag(.ColTag("UNIDCONTROLE"), "EDIT", False)
      .ColTag("VLUNIT") = SetTag(.ColTag("VLUNIT"), "EDIT", False)
      .ColTag("Total") = SetTag(.ColTag("Total"), "EDIT", False)
            
      .ColTag("QTDVENDA") = SetTag(.ColTag("QTDVENDA"), "EDITOPT", igTextEditNumberOnly)
      
      .ColTag("IDITEM") = SetTag(.ColTag("IDITEM"), "DEFAULT", "IDENTITY")
      .ColTag("IDPROD") = SetTag(.ColTag("IDPROD"), "DEFAULT", 1)
      .ColTag("UNIDCONTROLE") = SetTag(.ColTag("UNIDCONTROLE"), "DEFAULT", "")
      .ColTag("QTDVENDA") = SetTag(.ColTag("QTDVENDA"), "DEFAULT", 0)
      .ColTag("VLUNIT") = SetTag(.ColTag("VLUNIT"), "DEFAULT", ValBr(0))
      .ColTag("Total") = SetTag(.ColTag("Total"), "DEFAULT", ValBr(0))
      
      IncluiGrdLinhaVenda

   End With
End Sub
Private Sub MontaGrdPgto()
   Dim TbCond  As Object 'TB_CCONDPGTO
   Dim TbForma As Object 'TB_CFORMAPGTO
   Dim i       As Integer
   Dim Sql     As String
   Dim sAux    As String
   
   Set TbCond = CriarObjeto("BANCO_3R.TB_CCONDPGTO")
   Set TbCond.xDb = mvarSys.xDb
   
   Set TbForma = CriarObjeto("BANCO_3R.TB_CFORMAPGTO")
   Set TbForma.xDb = mvarSys.xDb
   
   
   With mvarMe.GrdPagamento
      .ShowControlsInAllCells = True
      .Header.DragCols = False
      .SilentValidation = True
      '.Header.ImageList = mvarMe.IlsIcons
      '.ImageList = mvarMe.IlsIcons
            
      With .Combos
         With .Add("IDFORMAPGTO")
            If TbForma.Pesquisar(Ch_Where:="ATIVO=1", Ch_ORDERBY:="IDFORMAPGTO") Then
               While Not TbForma.Rs.EOF
                  TbForma.Popula
                  sAux = UCase(TbForma.FORMAPGTO)
                  sAux = sAux & Space(15 - Len(sAux))
                  .AddItem sAux, CStr(TbForma.IDFORMAPGTO)

                  TbForma.Rs.MoveNext
               Wend
            End If
            Set TbForma = Nothing
            .AutoAdjustWidth
         End With
         With .Add("IDCONDPGTO")
            If TbCond.Pesquisar(Ch_Where:="IDCONDPGTO=1", Ch_ORDERBY:="IDCONDPGTO") Then
               While Not TbCond.Rs.EOF
                  TbCond.Popula
                  sAux = TbCond.CONDPGTO
                  sAux = sAux & Space(15 - Len(sAux))
                  .AddItem sAux, CStr(TbCond.IDCONDPGTO)

                  TbCond.Rs.MoveNext
               Wend
            End If
            Set TbCond = Nothing
            .AutoAdjustWidth
         End With
      End With
      
      With .AddCol(sKey:="IDPGTO", sHeader:="#", lWidth:=25, bVisible:=True)
         .eType = igCellText
         .sCtrlKey = "IDPGTO"
         
      End With
      With .AddCol(sKey:="IDFORMAPGTO", sHeader:="Forma", lWidth:=100, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDFORMAPGTO"
      End With
      With .AddCol(sKey:="IDCONDPGTO", sHeader:="Condição", lWidth:=100, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDCONDPGTO"
      End With
      With .AddCol(sKey:="VLPGTO", sHeader:="Valor", lWidth:=70, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight
         .sCtrlKey = "VLPGTO"
      End With
            
      For i = 1 To .ColCount
         .ColTag(i) = ""
      Next
            
      .ColTag("IDPGTO") = SetTag(.ColTag("IDPGTO"), "EDIT", False)
            
      '.ColTag("VLPGTO") = SetTag(.ColTag("VLPGTO"), "EDITOPT", igTextEditNumberOnly)
            
            
      IncluiGrdLinhaPgto

   End With
   
   Set TbCond = Nothing
   Set TbForma = Nothing
End Sub
Private Sub MontaGrdAtendimento()
   Dim TbCond  As Object 'TB_CCONDPGTO
   Dim TbForma As Object 'TB_CFORMAPGTO
   Dim i       As Integer
   Dim Sql     As String
   Dim sAux    As String
   
   Set TbCond = CriarObjeto("BANCO_3R.TB_CCONDPGTO")
   Set TbCond.xDb = mvarSys.xDb
   
   Set TbForma = CriarObjeto("BANCO_3R.TB_CFORMAPGTO")
   Set TbForma.xDb = mvarSys.xDb
   
   
   With mvarMe.GrdAtendimento
      .ShowControlsInAllCells = True
      .Header.DragCols = False
      .SilentValidation = True
      '.Header.ImageList = mvarMe.IlsIcons
      '.ImageList = mvarMe.IlsIcons
           
      With .AddCol(sKey:="IDATENDIMENTO", sHeader:="Nº", lWidth:=50, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft
         .sCtrlKey = "IDATENDIMENTO"
         
      End With
      With .AddCol(sKey:="TEL1", sHeader:="Telefone", lWidth:=80, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextCenter
         .sCtrlKey = "TEL1"
      End With
      With .AddCol(sKey:="IDCLIENTE", sHeader:="Cliente", lWidth:=100, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft
         .sCtrlKey = "IDCLIENTE"
      End With
      
      For i = 1 To .ColCount
         .ColTag(i) = ""
      Next
            
      .ColTag("IDATENDIMENTO") = SetTag(.ColTag("IDATENDIMENTO"), "EDIT", False)
      .ColTag("TEL1") = SetTag(.ColTag("TEL1"), "EDIT", False)
      .ColTag("IDCLIENTE") = SetTag(.ColTag("IDCLIENTE"), "EDIT", False)
            
      .ColTag("IDATENDIMENTO") = SetTag(.ColTag("IDATENDIMENTO"), "EDITOPT", igTextEditNumberOnly)
      .ColTag("TEL1") = SetTag(.ColTag("TEL1"), "EDITOPT", igTextEditNumberOnly)
            
            
      IncluiGrdLinhaAtendimento

   End With
   
   Set TbCond = Nothing
   Set TbForma = Nothing
End Sub
Private Sub IncluiGrdLinhaVenda()
   Dim lRow As Long
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim bAchou As Boolean
   Dim nColDefault As Integer
      
   nColDefault = 2
   With mvarMe.GrdVenda
      '.Header.Font.Size = 9
      .Header.Font.Name = mvarMe.Font.Name
      .Font.Size = .Header.Font.Size
      .Font.Name = .Header.Font.Name
      
      lRow = .RowCount
      If lRow > 0 Then .RemoveRow lRow
      If lRow > 1 Then .CellForeColor(lRow - 1, "IDPROD") = vbBlack
      
      'If Not (mvarIDATENDIMENTO = 0 And lRow = 0) Then
      If lRow > 0 Then
         .AddRow
         lRow = .RowCount
         .RowHeight(.RowCount) = 20
   
         For i = 1 To .ColCount
            If .CellType(lRow, .ColKey(i)) = igCellCombo Then
               If .Combos(.ColKey(i)).ListCount = 1 Then
                  .CellValue(lRow, .ColKey(i)) = .Combos(.ColKey(i)).ItemValue(0)
               ElseIf .Combos(.ColKey(i)).ListCount > 1 Then
                  If GetTag(.ColTag(.ColKey(i)), "DEFAULT", "") <> "" Then
                     If IsNumeric(GetTag(.ColTag(.ColKey(i)), "DEFAULT", "")) Then
                        .CellValue(lRow, .ColKey(i)) = Val(GetTag(.ColTag(.ColKey(i)), "DEFAULT", "0"))
                     Else
                        .CellValue(lRow, .ColKey(i)) = Val(GetTag(.ColTag(.ColKey(i)), "DEFAULT", ""))
                     End If
                     If lRow > 1 Then
                        For j = 0 To .Combos(.ColKey(i)).ListCount - 1
                           bAchou = False
                           If Not bAchou Then
                              .CellValue(lRow, .ColKey(i)) = .Combos(.ColKey(i)).ItemValue(j)
                           End If
                           For k = 1 To (lRow - 1)
                              bAchou = (.CellValue(lRow, .ColKey(i)) = .CellValue(k, .ColKey(i)))
                              If bAchou Then Exit For
                           Next
                        Next
                     End If
                     
                  End If
               End If
            Else
               If GetTag(.ColTag(.ColKey(i)), "DEFAULT", "") = "IDENTITY" Then
                  .CellValue(lRow, .ColKey(i)) = lRow
               Else
                  .CellValue(lRow, i) = GetTag(.ColTag(.ColKey(i)), "DEFAULT", "")
               End If
            End If

            If GetTag(.ColTag(i), "EDIT", True) = False Then
               .CellForeColor(lRow, i) = vbGrayText
            End If
         Next
         Call mvarMe_GrdVendaAfterCommitEdit(lRow, mvarMe.GrdVenda.ColIndex("IDPROD"))
         
         If Val(.CellValue(lRow, "IDPROD")) <> 0 And Val(.CellValue(lRow, "QTDVENDA")) <> 0 Then
            Call mvarMe_GrdVendaAfterCommitEdit(lRow, .ColIndex("IDPROD"))
         End If
      End If
      
      .AddRow
      For i = 1 To .ColCount
         .CellType(.RowCount, i) = igCellText
      Next
      For i = 1 To .ColCount
         If .ColVisible(i) Then
            .CellValue(.RowCount, i) = "Clique para incluir nova linha."
            .CellForeColor(.RowCount, i) = vbGrayText
            .CellTextFlags(.RowCount, i) = igTextNoClip Or igTextCenter
            Exit For
         End If
      Next
      
      
      .RowMode = (.RowCount = 1)
      If lRow <> 0 Then
         .SetCurCell .RowCount, 1
         .SetCurCell lRow, nColDefault
      End If
      If .Visible Then
         .SetFocus
      End If
      .Redraw = True
   End With
End Sub
Private Sub IncluiGrdLinhaPgto()
   Dim lRow As Long
   Dim i As Integer
   Dim nColDefault As Integer
   Dim nVlAPagar  As Currency
   Dim nVlPgto    As Currency
     
   If Not GetTag(mvarMe, "1VEZ", "1") = "1" Then
      If xVal(mvarMe.TxtVLVENDA.Text) - xVal(mvarMe.TxtVLDESC) <= xVal(mvarMe.TxtVLPGTO) Then
         If xVal(mvarMe.TxtVLPGTO) <> 0 Then
            Call ExibirAviso("Não há valor a pagar.")
            Exit Sub
         End If
      End If
   End If
   nColDefault = 2
   With mvarMe.GrdPagamento
      '.Header.Font.Size = 9
      .Header.Font.Name = mvarMe.Font.Name
      .Font.Size = .Header.Font.Size
      .Font.Name = .Header.Font.Name
      
      lRow = .RowCount
      If lRow > 0 Then .RemoveRow lRow
      If lRow > 1 Then .CellForeColor(lRow - 1, "IDPGTO") = vbBlack
      
      'If Not (mvarIDVENDA = 0 And lRow = 0) Then
      If lRow > 0 Then
         .AddRow
         lRow = .RowCount
         .RowHeight(.RowCount) = 20
   
         .CellType(lRow, "IDPGTO") = igCellText
         .CellValue(lRow, "IDPGTO") = lRow
         .CellForeColor(lRow, "IDPGTO") = vbGrayText
         
         .CellType(lRow, "IDCONDPGTO") = igCellCombo
         If .Combos("IDCONDPGTO").ListCount >= 1 Then
            .CellValue(lRow, "IDCONDPGTO") = .Combos("IDCONDPGTO").ItemValue(0)
         End If
   
         .CellType(lRow, "IDFORMAPGTO") = igCellCombo
         If .Combos("IDFORMAPGTO").ListCount >= 1 Then
            .CellValue(lRow, "IDFORMAPGTO") = .Combos("IDFORMAPGTO").ItemValue(0)
         End If
   
         .CellType(lRow, "VLPGTO") = igCellText
         .CellValue(lRow, "VLPGTO") = 0
         
         nVlPgto = 0
         For i = 1 To .RowCount
            nVlPgto = nVlPgto + xVal(.CellValue(i, "VLPGTO"))
         Next
         nVlAPagar = xVal(mvarMe.TxtVLVENDA.Text)
         nVlAPagar = nVlAPagar - xVal(mvarMe.TxtVLDESC.Text)
         nVlAPagar = nVlAPagar - nVlPgto
         .CellValue(lRow, "VLPGTO") = ValBr(nVlAPagar)
         
         Call PopulaValores
'         nVlPgto = 0
'         For i = 1 To .RowCount
'            nVlPgto = nVlPgto + xval(.CellValue(i, "VLPGTO"))
'         Next
'         mvarMe.TxtVLPGTO.Text = ValBr(nVlPgto)
   
         For i = 1 To .ColCount
            If GetTag(.ColTag(i), "EDIT", True) = False Then
               .CellForeColor(lRow, i) = vbGrayText
            End If
         Next
   
      End If
      .AddRow
     
      For i = 1 To .ColCount
         .CellType(.RowCount, i) = igCellText
      Next
      For i = 1 To .ColCount
         If .ColVisible(i) Then
            .CellValue(.RowCount, i) = "Clique para incluir nova linha."
            .CellForeColor(.RowCount, i) = vbGrayText
            .CellTextFlags(.RowCount, i) = igTextNoClip Or igTextCenter
            Exit For
         End If
      Next
      
      
      .RowMode = (.RowCount = 1)
      If lRow <> 0 Then
         .SetCurCell .RowCount, 1
         .SetCurCell lRow, nColDefault
      End If
      If .Visible Then
         .SetFocus
      End If
      .Redraw = True
   End With
End Sub
Private Sub IncluiGrdLinhaAtendimento(Optional pBuscaAtendimento As Boolean = False)
   Dim Sql     As String
   Dim MyAdd   As Object 'DsAdd
   Dim n       As Object
   Dim i       As Integer
   Dim MySel   As Collection
   
   Sql = "Select Distinct Right(REPLICATE('0'," + CStr(mvarSys.Propriedades("LENIDVENDA")) + ")+ Cast(A.IDATENDIMENTO as varchar)," + CStr(mvarSys.Propriedades("LENIDVENDA")) + ") [IDATENDIMENTO]"
   Sql = Sql & " , Convert(varchar, A.DTATEND, 103) [DTATEND], C.TEL1, C.NOME"
   Sql = Sql & " From OATENDIMENTO A"
   Sql = Sql & " Left Join OSESSAO S On A.IDLOJA=S.IDLOJA And A.IDATENDIMENTO=S.IDATENDIMENTO And S.IDTPSERVICO IN (Select Distinct IDTPSERVICO From OTRATAMENTO_PROD)"
   Sql = Sql & " Left Join OATENDIMENTO_PRODUTO P On P.IDLOJA=A.IDLOJA And P.IDATENDIMENTO=A.IDATENDIMENTO"
   Sql = Sql & " Left Join OCLIENTE C On A.IDLOJA=C.IDLOJA And A.IDCLIENTE=C.IDCLIENTE"
   Sql = Sql & " Left Join OATENDIMENTO_VENDA V On A.IDLOJA=V.IDLOJA And A.IDATENDIMENTO=V.IDATENDIMENTO"
   Sql = Sql & " Where V.IDVENDA Is Null"
   Sql = Sql & " And (Not S.IDTPSERVICO Is Null Or Not P.IDPROD Is Null)"
   If mvarMe.GrdAtendimento.RowCount > 1 Then
      Sql = Sql & " And A.IDATENDIMENTO Not In ( "
      For i = 1 To mvarMe.GrdAtendimento.RowCount - 1
         Sql = Sql & IIf(i = 1, "", ", ")
         Sql = Sql & Val(mvarMe.GrdAtendimento.CellValue(i, "IDATENDIMENTO"))
      Next
      Sql = Sql & " )"
   End If
   Sql = Sql & " Order By IDATENDIMENTO"
      
'******************************************
'******************************************
   Dim lRow As Long
   Dim nColDefault As Integer
   
   nColDefault = 2
   With mvarMe.GrdAtendimento
      '.Header.Font.Size = 9
      .Header.Font.Name = mvarMe.Font.Name
      .Font.Size = .Header.Font.Size
      .Font.Name = .Header.Font.Name
      
      lRow = .RowCount
      If lRow > 0 Then .RemoveRow lRow
      If lRow > 1 Then .CellForeColor(lRow - 1, "IDATENDIMENTO") = vbBlack
      
      'If Not (mvarIDATENDIMENTO = 0 And lRow = 0) Then
      If lRow > 0 Then
         .AddRow
         lRow = .RowCount
         .RowHeight(.RowCount) = 15
   
         For i = 1 To .ColCount
            If .CellType(lRow, .ColKey(i)) = igCellCombo Then
               If .Combos(.ColKey(i)).ListCount = 1 Then
                  .CellValue(lRow, .ColKey(i)) = .Combos(.ColKey(i)).ItemValue(0)
               ElseIf .Combos(.ColKey(i)).ListCount > 1 Then
                  .CellValue(lRow, .ColKey(i)) = GetTag(.ColTag(.ColKey(i)), "DEFAULT", "")
               End If
            Else
               If GetTag(.ColTag(.ColKey(i)), "DEFAULT", "") = "IDENTITY" Then
                  .CellValue(lRow, .ColKey(i)) = lRow
               Else
                  .CellValue(lRow, i) = GetTag(.ColTag(.ColKey(i)), "DEFAULT", "")
               End If
            End If
         Next
      End If
      
      If .RowCount >= 1 And pBuscaAtendimento Then
         'Set MyAdd = New DsAdd
         Set MyAdd = CriarObjeto("XActive.DsAdd")
         With MyAdd
            .xDb = mvarSys.xDb
            .Caption = "Atendimentos Sem Venda."
            .Chaves.Add "IDATENDIMENTO", "IDATENDIMENTO"
            .Chaves.Add "TEL1", "TEL1"
            .Chaves.Add "NOME", "NOME"
            With .TITULO
               Call .Add("IDATENDIMENTO", "Nº Atend.", 8, 0) 'eAlin.Esquerda)
               Call .Add("DTATEND", "Data", 10, 0) 'eAlin.Esquerda)
               Call .Add("TEL1", "Telefone", 10, 2) 'eAlin.Centro)
               Call .Add("NOME", "Nome", 30, 1) 'eAlin.Esquerda)
            End With
            .CheckBox = True
            .QUERY = Sql
      
            .Show
            Set MySel = .Selecao
         End With
         
         .RemoveRow lRow
         For Each n In MySel
            .AddRow
            .CellValue(.RowCount, "IDATENDIMENTO") = n("IDATENDIMENTO")
            .CellValue(.RowCount, "TEL1") = n("TEL1")
            .CellValue(.RowCount, "IDCLIENTE") = n("NOME")
            
            For i = 1 To .ColCount
               If GetTag(.ColTag(i), "EDIT", True) = False Then
                  .CellForeColor(.RowCount, i) = vbGrayText
               End If
            Next
            
         Next
         Call PopulaGRDPRODUTO
         Call CalculaVenda
         Call PopulaValores
      End If
      
      .AddRow
     
      For i = 1 To .ColCount
         .CellType(.RowCount, i) = igCellText
      Next
      For i = 1 To .ColCount
         If .ColVisible(i) Then
            .CellValue(.RowCount, i) = "Clique para incluir nova linha."
            .CellForeColor(.RowCount, i) = vbGrayText
            .CellTextFlags(.RowCount, i) = igTextNoClip Or igTextCenter
            Exit For
         End If
      Next
      
      .RowMode = (.RowCount = 1)
      If lRow <> 0 Then
         .SetCurCell .RowCount, 1
         .SetCurCell lRow, nColDefault
      End If
      If .Visible Then
         .SetFocus
      End If
      .Redraw = True
   End With
End Sub
Private Sub PopulaCliente(Optional pIDCLIENTE As Long, Optional pTelefone As String = "")
   Dim TbCli   As Object 'TB_OCLIENTE
   Dim bAchou  As Boolean
   
   Set TbCli = CriarObjeto("BANCO_3R.TB_OCLIENTE")
   With TbCli
      Set .xDb = mvarSys.xDb
      If IsMissing(pIDCLIENTE) Then
         bAchou = .Pesquisar(Ch_Where:="TEL1='" & pTelefone & "'")
         If Not bAchou Then bAchou = .Pesquisar(Ch_Where:="TEL2='" & pTelefone & "'")
      Else
         bAchou = .Pesquisar(Ch_IDCLIENTE:=pIDCLIENTE)
      End If
      If bAchou Then
         mvarIDCLIENTE = .IDCLIENTE
         Call SetTag(mvarMe.TxtNOME, "IDCLIENTE", .IDCLIENTE)
         mvarMe.TxtNOME.Text = .NOME
         mvarMe.TxtTEL1.Text = IIf(.TEL1 = "", .TEL2, .TEL1)
      Else
         mvarIDCLIENTE = 0
         Call SetTag(mvarMe.TxtNOME, "IDCLIENTE", "")
         mvarMe.TxtNOME.Text = ""
         mvarMe.TxtTEL1.Text = ""
      End If
   End With
   Set TbCli = Nothing
End Sub
Private Sub mvarMe_TxtIDVENDALostFocus()
   If Val(mvarMe.TxtIDVENDA.Text) = 0 Then
      mvarMe.TxtIDVENDA.Text = ""
   Else
      mvarMe.TxtIDVENDA.Text = StrZero(mvarMe.TxtIDVENDA.Text, mvarSys.Propriedades("LENIDVENDA"))
   End If
End Sub
Private Function GetProdutoVenda(pIDTPSERVICO As Long, pIDTPTRATAMENTO As Long, pIDAREA As Long) As Long
   Dim Sql As String
   Dim MyRs As Object
   
   Sql = ""
   Sql = Sql & "Select IDTPSERVICO, IDTPTRATAMENTO, IDAREA, IDPROD"
   Sql = Sql & " From OTRATAMENTO_PROD"
   Sql = Sql & " Where ( IDTPSERVICO is Null Or IDTPSERVICO =" & pIDTPSERVICO & " )"
   Sql = Sql & " And ( IDTPTRATAMENTO is Null Or IDTPTRATAMENTO =" & pIDTPTRATAMENTO & " )"
   Sql = Sql & " And ( IDAREA is Null Or IDAREA =" & pIDAREA & " )"
   Sql = Sql & " Order By IDTPSERVICO DESC, IDTPTRATAMENTO DESC, IDAREA DESC, IDPROD"
   
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      GetProdutoVenda = Val(MyRs("IDPROD") & "")
   End If
End Function
Private Function CalculaVenda() As Long
   Dim i As Integer
   
   With mvarMe.GrdVenda
      For i = 1 To .RowCount - 1
         CalculaVenda = xVal(CalculaVenda) + xVal(.CellValue(i, "Total"))
      Next
   End With
   mvarMe.TxtVLVENDA.Text = ValBr(CalculaVenda)
End Function
Private Sub mvarMe_TxtTEL1LostFocus()
   Dim NG_Cal As Object 'NG_Calendario
   Dim sTag As String
   
   Set NG_Cal = CriarObjeto("Calendario3R.NG_Calendario") 'New NG_Calendario
   With NG_Cal
      .IDLOJA = mvarIDLOJA
      Set .Sys = mvarSys
      sTag = .MProcuraCliente(mvarMe.TxtTEL1.Text, mvarMe.TxtNOME.Text, mvarIDCLIENTE)
      If sTag <> "" Then
         mvarIDCLIENTE = GetTag(sTag, "IDCLIENTE", "0")
         mvarMe.TxtNOME.Text = GetTag(sTag, "NOME", "")
         mvarMe.TxtTEL1.Text = GetTag(sTag, "TEL1", "")
         If mvarMe.GrdVenda.Enabled Then
            mvarMe.GrdVenda.SetFocus
         End If
         
         Call SetTag(mvarMe.TxtNOME, "IDCLIENTE", mvarIDCLIENTE)
      End If
   End With
End Sub
Private Sub mvarMe_TxtVLDESCChange()
   PopulaValores
End Sub
Private Sub mvarMe_TxtVLPGTOChange()
   PopulaValores
End Sub

