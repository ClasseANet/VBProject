VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Venda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe    As FrmVenda
Attribute mvarMe.VB_VarHelpID = -1

Private Const BTN_SIZE = 14
Private Const ROW_COUNT = 50

' these vars used to store the row and col of the cell with pushed button:
Private m_lBtnRow As Long
Private m_lBtnCol As Long
Private m_lBtnLeft As Long
Private m_lBtnTop As Long
Private m_lBtnRight As Long
Private m_lBtnBottom As Long
Private arrCellText(1 To ROW_COUNT) As String


Dim nLenIDVenda   As Integer
Dim bVendaFechada As Boolean
Dim dToday        As Date
Dim bExibeFatura  As Boolean
Dim bExistePromo  As Boolean
Dim m_oFont       As StdFont
Dim CItensNew     As Collection

Private mvarSys            As Object
Private mvarIDLOJA         As Integer
Private mvarIDATENDIMENTO  As Long
Private mvarIDCLIENTE      As Long
Private mvarIDFUNCIONARIO  As Long
Private mvarIDVENDA        As Long
Private mvarIDFATURA       As Long
Private mvarIDFATVENDA     As Long
Private mvarCalControl     As CalendarControl
Private mvarCalEvent       As CalendarEvent
Private mvarTlVendas       As TL_VENDAS
Private mvarPacotes        As Collection
Public Property Set TlVendas(ByRef vData As Object)
   Set mvarTlVendas = vData
End Property
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
Public Property Let IDFATURA(ByVal vData As Long)
    mvarIDFATURA = vData
End Property
Public Property Get IDFATURA() As Long
   IDFATURA = mvarIDFATURA
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
Public Property Let IDFUNCIONARIO(ByVal vData As Long)
    mvarIDFUNCIONARIO = vData
End Property
Public Property Get IDFUNCIONARIO() As Long
   IDFUNCIONARIO = mvarIDFUNCIONARIO
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
   dToday = CDate(Format(mvarSys.xDb.SysDate, "DD/MM/YYYY"))
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmVenda
   Set mvarPacotes = New Collection
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
   Set mvarPacotes = Nothing
End Sub
Public Sub Show(Optional pMode As FormShowConstants = vbModal)
'   If mvarIDLOJA = 0 Then
'      Call ExibirAviso("Código da Loja Inválido! (""" & mvarIDLOJA & """)")
'      Exit Sub
'   End If
   
   mvarMe.Show pMode
End Sub
Private Sub mvarMe_CmdCancelClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdChaveClick()
   If SenhaMestre(mvarSys) Then
      DoEvents
      bVendaFechada = False
      Call SetTag(mvarMe, "1VEZ", "1")
      'Call LimparTela
      Call PopulaTela
      Call SetTag(mvarMe, "1VEZ", "0")
      mvarMe.CmdFechar.Enabled = True
      mvarMe.CmdExcluir.Visible = True
      mvarMe.CmbDTVENDA.Enabled = True
   End If
End Sub

Private Sub mvarMe_CmdDescClick()
   Dim TlDesc As TL_CupomVenda
   Set TlDesc = New TL_CupomVenda
   With TlDesc
      Set .Sys = mvarSys
      .Show
   End With
   Set TlDesc = Nothing
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
   Dim TbPaco As Object 'TB_CPACOTE
   Dim TbCupo As Object 'TB_CCUPOM
   Dim TbNF   As Object 'TB_FNOTAFISCAL
   
   Dim TbMovEst As Object 'TB_SMOVEST
   Dim TbEstProd As Object 'TB_SESTPROD
   
   sMsg = "Esta operação irá excluir definitivamente a venda '" & mvarMe.TxtIDVENDA.Text & "'"
   sMsg = sMsg & vbNewLine & vbNewLine
   sMsg = sMsg & "Deseja continuar?"
   If ExibirPergunta(sMsg, , False) = vbYes Then
      'If InputBoxPassword("Informe a senha.") = mvarSys.Propriedades("SENHAMESTRE") Then
         Set Queries = New Collection
         
         Set TbVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
         With TbVenda
            Set .xDb = mvarSys.xDb
            If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA) Then
               Set TbItem = CriarObjeto("BANCO_3R.TB_CITENSVENDA")
               Set TbPgto = CriarObjeto("BANCO_3R.TB_CPGTOSVENDA")
               Set TbAten = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_VENDA")
               Set TbMovEst = CriarObjeto("BANCO_3R.TB_SMOVEST")
               Set TbEstProd = CriarObjeto("BANCO_3R.TB_SESTPROD")
               Set TbNF = CriarObjeto("BANCO_3R.TB_FNOTAFISCAL")
               Set TbPaco = CriarObjeto("BANCO_3R.TB_CPACOTE")
               Set TbCupo = CriarObjeto("BANCO_3R.TB_CCUPOM")
               Set TbItem.xDb = mvarSys.xDb
               Set TbPgto.xDb = mvarSys.xDb
               Set TbAten.xDb = mvarSys.xDb
               Set TbMovEst.xDb = mvarSys.xDb
               Set TbEstProd.xDb = mvarSys.xDb
               Set TbNF.xDb = mvarSys.xDb
               Set TbPaco.xDb = mvarSys.xDb
               Set TbCupo.xDb = mvarSys.xDb
                              
               Queries.Add TbItem.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA)
               Queries.Add TbPgto.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA)
               Queries.Add TbAten.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA)
               Queries.Add TbPaco.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA)
               Queries.Add TbCupo.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDVENDA=" & mvarIDVENDA)
               
               If TbMovEst.Pesquisar(Ch_Where:="IDLOJA=" & mvarIDLOJA & " And IDVENDA=" & mvarIDVENDA) Then
                  While Not TbMovEst.Rs.EOF
                     TbMovEst.Popula
                     Queries.Add TbMovEst.QryDelete(Ch_IDMOV:=TbMovEst.IDMOV)
                     TbMovEst.Rs.MoveNext
                  Wend
               End If
               
               '* Recuperar Qtd Estoque
               Sql = "Select IDPROD, Sum(QTDITEM) [QTD]"
               Sql = Sql & " From SMOVEST"
               Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
               Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
               Sql = Sql & " Group And IDVENDA=" & SqlNum(mvarIDVENDA)
               If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
                  While Not MyRs.EOF
                     Sql = "Update SESTPROD Set SLDATUAL , SLDDISPONIVEL"
                     Sql = Sql & " Set SLDATUAL=SLDATUAL+" & MyRs("QTD")
                     Sql = Sql & ", SLDDISPONIVEL=SLDDISPONIVEL+" & MyRs("QTD")
                     Sql = Sql & ", ALTERSTAMP=1"
                     Sql = Sql & ", TIMESTAMP=GetDate()"
                     Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
                     Sql = Sql & " And IDPROD=" & MyRs("IDPROD")
                     Sql = Sql & ";"
                     MyRs.MoveNext
                  Wend
               End If
                              
               '* Todos os Lançamentos da Venda
               Sql = "Delete FLAN "
               Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
               Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
               Sql = Sql & " Insert Into DELETEDROWS (QUERY, SITQUERY, TIMESTAMP, ALTERSTAMP)"
               Sql = Sql & "  Output Inserted.*"
               Sql = Sql & "  Select 'Delete From FLAN"
               Sql = Sql & "  Where IDLOJA = " & SqlNum(mvarIDLOJA)
               Sql = Sql & "  And IDVENDA=" & SqlNum(mvarIDVENDA) & " ;'"
               Sql = Sql & ", 0"
               Sql = Sql & ", " & "GetDate()" & vbNewLine
               Sql = Sql & ", " & "1" & vbNewLine
               Sql = Sql & " ;"
               Queries.Add Sql
                              
                                             
               nRecibo = xVal(Trim(Mid(mvarMe.CmdRecibo.Caption, Len("Recibo "))))
               Call ExcluirRPS(mvarIDLOJA, nRecibo, Queries)
                              
               Sql = "Update FFATURA "
               Sql = Sql & " Set SITFATURA=0"
               Sql = Sql & ", IDVENDA =Null"
               Sql = Sql & ", ALTERSTAMP=1"
               Sql = Sql & ", TIMESTAMP=GetDate()"
               Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
               Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
               Sql = Sql & ";"
               Queries.Add Sql
               
               Queries.Add .QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA)
               
               '*************
               '* Se não possui Nota Fiscal
               If Not TbNF.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDVENDA=" & mvarIDVENDA) Then
                  Sql = "Update OSESSAO "
                  Sql = Sql & " Set IDVENDA=0"
                  Sql = Sql & " ,IDITEM=0"
                  Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
                  Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
                  Sql = Sql & ";"
                  Queries.Add Sql
                  
                  Sql = "Update OSESSAO "
                  Sql = Sql & " Set IDVENDA=IDVENDA-1 "
                  Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
                  Sql = Sql & " And IDVENDA>" & SqlNum(mvarIDVENDA)
                  Sql = Sql & ";"
                  Queries.Add Sql
                  
                  Sql = "Update CVENDA"
                  Sql = Sql & " Set IDVENDA=IDVENDA-1"
                  Sql = Sql & ", ALTERSTAMP=1"
                  Sql = Sql & ", TIMESTAMP=GetDate()"
                  Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
                  Sql = Sql & " And IDVENDA>" & SqlNum(mvarIDVENDA)
                  Sql = Sql & ";"
                  Queries.Add Sql
               
                  If TbVenda.CODVENDA <> 0 Then
                     Sql = "Update CVENDA"
                     Sql = Sql & " Set CODVENDA=CODVENDA-1"
                     Sql = Sql & ", ALTERSTAMP=1"
                     Sql = Sql & ", TIMESTAMP=GetDate()"
                     Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
                     Sql = Sql & " And IDVENDA>" & SqlNum(mvarIDVENDA)
                     Sql = Sql & " And CODVENDA>2"
                     Sql = Sql & ";"
                     Queries.Add Sql
                  End If
                  
                  Sql = "Update FLAN "
                  Sql = Sql & " Set IDVENDA=IDVENDA-1"
                  Sql = Sql & ", ALTERSTAMP=1"
                  Sql = Sql & ", TIMESTAMP=GetDate()"
                  Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
                  Sql = Sql & " And IDVENDA>=" & SqlNum(mvarIDVENDA)
                  Sql = Sql & ";"
                  Queries.Add Sql

                  Sql = "Update FLAN "
                  Sql = Sql & " Set NDOC=RIGHT('000000' + CAST(IDVENDA AS VARCHAR),6)"
                  Sql = Sql & ", ALTERSTAMP=1"
                  Sql = Sql & ", TIMESTAMP=GetDate()"
                  Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
                  Sql = Sql & " And IDVENDA>" & SqlNum(mvarIDVENDA)
                  Sql = Sql & ";"
                  Queries.Add Sql
               
                  Sql = "Update SMOVEST "
                  Sql = Sql & " Set IDVENDA=IDVENDA-1"
                  Sql = Sql & ", ALTERSTAMP=1"
                  Sql = Sql & ", TIMESTAMP=GetDate()"
                  Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
                  Sql = Sql & " And IDVENDA>" & SqlNum(mvarIDVENDA)
                  Sql = Sql & ";"
                  Queries.Add Sql
               
                  Sql = "Update SMOVEST "
                  Sql = Sql & " Set NUMDOC=RIGHT('000000' + CAST(IDVENDA AS VARCHAR),6)"
                  Sql = Sql & ", ALTERSTAMP=1"
                  Sql = Sql & ", TIMESTAMP=GetDate()"
                  Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
                  Sql = Sql & " And IDVENDA>" & SqlNum(mvarIDVENDA)
                  Sql = Sql & ";"
                  Queries.Add Sql
                              
                  Sql = "Update FFATURA"
                  Sql = Sql & " Set IDVENDA=IDVENDA-1"
                  Sql = Sql & ", ALTERSTAMP=1"
                  Sql = Sql & ", TIMESTAMP=GetDate()"
                  Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
                  Sql = Sql & " And IDVENDA>" & SqlNum(mvarIDVENDA)
                  Sql = Sql & ";"
                  Queries.Add Sql
               
                  Sql = "Update FFATURA"
                  Sql = Sql & " Set IDVENDAORIGEM=IDVENDAORIGEM-1"
                  Sql = Sql & ", ALTERSTAMP=1"
                  Sql = Sql & ", TIMESTAMP=GetDate()"
                  Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
                  Sql = Sql & " And IDVENDAORIGEM>" & SqlNum(mvarIDVENDA)
                  Sql = Sql & ";"
                  Queries.Add Sql
                  
                  Sql = "Update CPACOTE"
                  Sql = Sql & " Set IDVENDA=IDVENDA-1"
                  Sql = Sql & ", ALTERSTAMP=1"
                  Sql = Sql & ", TIMESTAMP=GetDate()"
                  Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
                  Sql = Sql & " And IDVENDA>=" & SqlNum(mvarIDVENDA)
                  Sql = Sql & ";"
                  'Queries.Add Sql  '*Cascade
               
                  Sql = "Update CCUPOM"
                  Sql = Sql & " Set IDVENDA=IDVENDA-1"
                  Sql = Sql & ", ALTERSTAMP=1"
                  Sql = Sql & ", TIMESTAMP=GetDate()"
                  Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
                  Sql = Sql & " And IDVENDA>=" & SqlNum(mvarIDVENDA)
                  Sql = Sql & ";"
                  Queries.Add Sql
               
               Else
                  Call ExibirAviso("Venda possui Nota Fiscal e não pode ser apagada.")
                  Exit Sub
               End If
               
               Call SalvarEvento(True)
               
               Dim i As Integer
               For i = 1 To Queries.Count
                  Debug.Print Queries(i)
               Next
               If mvarSys.xDb.Executa(Queries) Then
                  mvarIDVENDA = 0
                  Call ExibirResultado(mvarSys, True)
                  If Not mvarTlVendas Is Nothing Then
                     Call mvarTlVendas.Refresh
                  End If
                  Unload mvarMe
               Else
                  Call SalvarEvento
                  Call ExibirResultado(mvarSys, False)
               End If
            End If
         End With
      'Else
      '   Call ExibirAviso("Esta senha não permite a operação.")
      'End If
   End If
   Set TbItem = Nothing
   Set TbPgto = Nothing
   Set TbAten = Nothing
   Set TbPaco = Nothing
   Set TbCupo = Nothing
End Sub
Public Sub ExcluirRPS(pIDLOJA As Integer, pRecibo As Long, ByRef pQueries As Collection, Optional pIDVENDA As Long)
   Dim TbReci As Object 'TB_FRECIBO
   Dim Sql As String
   Dim MyRs As Object
   
   If pQueries Is Nothing Then Set pQueries = New Collection
   
   Set TbReci = CriarObjeto("BANCO_3R.TB_FRECIBO")
   Set TbReci.xDb = mvarSys.xDb
   
   If pIDVENDA <> 0 And pRecibo = 0 Then
      Sql = " IDVENDA=" & SqlNum(mvarIDVENDA)
      Sql = Sql & " And SERIE=1"
      If TbReci.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=Sql) Then
         pRecibo = TbReci.IDRECIBO
      End If
   End If
   If TbReci.Pesquisar(Ch_IDLOJA:=pIDLOJA, Ch_IDRECIBO:=pRecibo, Ch_SERIE:=1) Then
      Sql = "Select Max(IDRECIBO)"
      Sql = Sql & " From FRECIBO "
      Sql = Sql & " Where IDLOTE=0 "
      Sql = Sql & " And SERIE=1"
      Sql = Sql & " And IDLOJA=" & SqlNum(pIDLOJA)
      If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
         If pRecibo = xVal(MyRs(0) & "") Then
            pQueries.Add TbReci.QryDelete(Ch_IDLOJA:=pIDLOJA, Ch_IDRECIBO:=pRecibo, Ch_SERIE:=1)
         Else
            If TbReci.IDVENDA = mvarIDVENDA Then
               pQueries.Add TbReci.QryDelete(Ch_IDLOJA:=pIDLOJA, Ch_IDRECIBO:=pRecibo, Ch_SERIE:=1)
               Sql = "Select Count(*)"
               Sql = Sql & " From FRECIBO"
               Sql = Sql & " Where IDLOTE<>0 "
               Sql = Sql & " And SERIE=1"
               Sql = Sql & " And IDLOJA=" & SqlNum(pIDLOJA)
               Sql = Sql & " And IDRECIBO>" & SqlNum(pRecibo)
               If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
                  If xVal(MyRs(0) & "") = 0 Then
                     Sql = "Update FRECIBO"
                     Sql = Sql & " Set IDRECIBO=IDRECIBO-1"
                     Sql = Sql & ", ALTERSTAMP=1"
                     Sql = Sql & ", TIMESTAMP=GetDate()"
                     Sql = Sql & " Where SERIE=1"
                     Sql = Sql & " And IDLOJA=" & SqlNum(pIDLOJA)
                     Sql = Sql & " And IDRECIBO>" & SqlNum(pRecibo)
                     Sql = Sql & ";"
                     pQueries.Add Sql
                  End If
               End If
            Else
               TbReci.Status = 2
               pQueries.Add TbReci.QryUpdate
            End If
         End If
      End If
   End If
   Set TbReci = Nothing
End Sub
Private Sub mvarMe_CmdFaturaClick()
   Dim nId As Long
   nId = xVal(mvarMe.CmdFatura.Caption)
   If nId > 0 Then
      Call ExibirFatura(nId)
   Else
      Call ExibirAviso("Fatura Inválida.")
   End If
End Sub

Private Sub mvarMe_CmdFaturasClick()
   Dim nId As Long
   nId = xVal(mvarMe.CmdFatura.Caption)
   If nId > 0 Then
      Call ExibirFatura(nId)
   Else
      Call ExibirAviso("Fatura Inválida.")
   End If
End Sub

Private Sub mvarMe_CmdFecharClick()
   Dim bAux As Boolean
   bAux = mvarMe.CmdOk.Enabled
   mvarMe.CmdOk.Enabled = False
   DoEvents
   If FecharVenda Then
      Unload mvarMe
   Else
      mvarMe.CmdOk.Enabled = bAux
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
   Dim MyRsAj  As Object
   Dim bNovo   As Boolean
   Dim nTaxa   As Double
   Dim nFatura As Long
   Dim nValor As Currency
   
   If Not VerificaFechamento() Then Exit Function
   
   If bExibeFatura Then
      nFatura = 0
      nValor = 0
      Sql = "Select isNull(Sum(VALOR),0)-" & SqlNum(CStr(xVal(UnFormat(mvarMe.TxtVLTROCO.Text)) / 100)) & " [VALOR]"
      Sql = Sql & " From FFATURA "
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And IDVENDAORIGEM=" & mvarIDVENDA
      If mvarIDVENDA <= 0 Then
         nValor = -1 * xVal(UnFormat(mvarMe.TxtVLTROCO.Text)) / 100
      Else
         If mvarSys.xDb.AbreTabela(Sql) Then
            nValor = xVal(mvarSys.xDb.RsAux("VALOR") & "")
         End If
      End If
      If nValor < 0 Then
         If mvarIDVENDA < 0 Then
            If Not SalvarVenda() Then Exit Function
         End If
         Call ExibirFatura(nFatura, pIDVENDAORIGEM:=mvarIDVENDA, pIDATENDIMENTO:=mvarIDATENDIMENTO, pIDCLIENTE:=mvarIDCLIENTE, pVLFATURA:=nValor * -1)
         If nFatura = 0 Then Exit Function
         Sql = "Select Sum(VALOR)-" & SqlNum(CStr(xVal(UnFormat(mvarMe.TxtVLTROCO.Text)) / 100)) & " [VALOR]"
         Sql = Sql & " From FFATURA "
         Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
         Sql = Sql & " And IDVENDAORIGEM=" & mvarIDVENDA
         If mvarSys.xDb.AbreTabela(Sql) Then
            nValor = xVal(mvarSys.xDb.RsAux("VALOR") & "")
            While nValor < 0
               If vbYes = ExibirPergunta("Fatura insuficiente para cobrir Saldo." & vbNewLine & vbNewLine & "Deseja gerar outra Fatura de R$" & ValBr(nValor * -1) & " ?") Then
                  Call ExibirFatura(0, mvarIDVENDA, mvarIDATENDIMENTO, mvarIDCLIENTE, nValor * -1)
                  If mvarSys.xDb.AbreTabela(Sql) Then
                     nValor = mvarSys.xDb.RsAux("VALOR")
                  End If
               Else
                  nValor = 0
                  Call ExibirInformacao("Pagamento inválido.")
                  Exit Function
               End If
            Wend
         End If
      End If
   End If
   
   If Not SalvarVenda() Then Exit Function
   
   Set Queries = New Collection
   
   '******************
   '* Fechar Venda
   '******************
   Set TbVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
   With TbVenda
      .xDb = mvarSys.xDb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA) Then
         If (xVal((.VLPGTO - .VLTROCO)) - xVal(.VLVENDA)) < 0 Then
            If .IDFATURA = 0 And mvarIDFATVENDA = 0 Then
               .IDFATURA = 1
            End If
            If .IDFATURA = 1 And mvarIDFATVENDA <> 0 Then
               .IDFATURA = mvarIDFATVENDA
            End If
         Else
            .IDFATURA = 0
         End If
         .SITVENDA = gSITVENDA_Fechada
         If .isDirt Then
            Queries.Add .QrySave
         End If
      End If
   End With
        
   '******************
   '* Lançar Pagamento
   '******************
   Set TbConta = CriarObjeto("BANCO_3R.TB_FCCORRENTE")
   TbConta.xDb = mvarSys.xDb
   
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
            Sql = " IDVENDA = " & SqlNum(mvarIDVENDA)
            Sql = Sql & " And IDPGTO = " & SqlNum(TbPgto.IDPGTO)
            Sql = Sql & " And NPARCELA = " & SqlNum(iParcela)
                                          
            Set TbFLAN = CriarObjeto("BANCO_3R.TB_FLAN")
            With TbFLAN
               .xDb = mvarSys.xDb
               bNovo = Not TbFLAN.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=Sql)
               
               '************
               '* Verifica se existe lançamento duplo. Não pode existir 2 lançamentos para o mesmo "IDPGTO" da venda.
               If mvarSys.xDb.RsAux.RecordCount > 1 Then
                  Sql = " IDVENDA = " & SqlNum(mvarIDVENDA)
                  Sql = Sql & " And IDPGTO = " & SqlNum(TbPgto.IDPGTO)
                  Sql = Sql & " And IDCONTA <> " & SqlNum(TbConta.IDCONTA)
                  Call mvarSys.xDb.Executa(TbFLAN.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=Sql))
               End If
               
               If bNovo Then
                  .CODSIT = ""
                  .DTCADASTRO = mvarSys.xDb.SysDate(3)
               End If
               
               .IDLOJA = mvarIDLOJA
               .IDCONTA = TbConta.IDCONTA
               
               .IDVENDA = mvarIDVENDA
               .IDPGTO = TbPgto.IDPGTO
               
               .NDOC = StrZero(TbPgto.IDVENDA, mvarSys.Propriedades("LENIDVENDA"))
               If TbPgto.IDFORMAPGTO = 1 Then '* Dinheiro
                  .Valor = TbPgto.VLPGTO - IIf(TbPgto.IDFORMAPGTO = 1, TbVenda.VLTROCO, 0)
               ElseIf TbPgto.IDFORMAPGTO = 2 Or TbPgto.IDFORMAPGTO = 3 Then '* Débito ou Crédito
                  nTaxa = 0
                  Sql = "Select TXSERV, TXPARC "
                  Sql = Sql & " From CFORMAPGTO"
                  Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
                  Sql = Sql & " And IDFORMAPGTO=" & TbPgto.IDFORMAPGTO
                  If mvarSys.xDb.AbreTabela(Sql) Then
                     If TbPgto.IDCONDPGTO <= 2 Then
                        nTaxa = xVal(mvarSys.xDb.RsAux("TXSERV") & "")
                     Else
                        nTaxa = xVal(mvarSys.xDb.RsAux("TXPARC") & "")
                     End If
                  End If
                  .Valor = TbPgto.VLPGTO - ValBr(TbPgto.VLPGTO * nTaxa / 100, 2, 2)
               Else
                  .Valor = TbPgto.VLPGTO
               End If
               
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
               .HISTORICO = mvarMe.TxtNOME.Text
   '            .SALDO = SqlNum(0, True)
   'XXX           .SALDO = CalcularSaldo
               
               '.DTCONTAB = ""
               '.DTBAIXA = ""
               .FLGDELETE = 0
               Call DefineParcela(TbFLAN, TbVenda.DTVENDA, TbPgto.IDFORMAPGTO, TbPgto.IDCONDPGTO, .Valor, iParcela)
               If TbPgto.IDFORMAPGTO = 3 And TbPgto.IDCONDPGTO = 1 Then
                  TbPgto.IDCONDPGTO = 2
               End If
               If .isDirt Then
                  'Queries.Add .QryDelete
                  Queries.Add .QrySave
               End If
            End With
         Next
         TbPgto.Rs.MoveNext
      Wend
   End If
      
   '******************
   '* Baixar Estoque
   '******************
   Sql = "Select V.IDLOJA, V.IDVENDA, V.DTVENDA, V.IDCLIENTE, I.IDITEM"
   Sql = Sql & ", I.IDPROD, I.QTDVENDA, I.UNIDCONTROLE"
   Sql = Sql & " From CVENDA V "
   Sql = Sql & " Join CITENSVENDA I On V.IDLOJA=I.IDLOJA And V.IDVENDA=I.IDVENDA "
   Sql = Sql & " Join SPRODUTO P    On I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD"
   Sql = Sql & " Where P.SITPROD='A'"
   Sql = Sql & " And P.EVENDA=1"
   Sql = Sql & " And P.ESERVICO=0"
   Sql = Sql & " And V.IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And V.IDVENDA=" & mvarIDVENDA
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      Sql = "Select L.IDEST"
      Sql = Sql & " From SEST_LOJA L"
      Sql = Sql & " Join SESTOQUE E On L.IDLOJA=E.IDLOJA And L.IDEST=E.IDEST  "
      Sql = Sql & " Where E.ATIVO=1"
      Sql = Sql & " And L.IDLOJA=" & mvarIDLOJA
      If mvarSys.xDb.AbreTabela(Sql) Then
         nIDEST = Val(mvarSys.xDb.RsAux("IDEST") & "")
      End If
      If nIDEST > 0 Then
         While Not MyRs.EOF
            Set TbEstP = CriarObjeto("BANCO_3R.TB_SESTPROD")
            Set TbEstP.xDb = mvarSys.xDb
            
            Call TbEstP.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEST:=nIDEST, Ch_IDPROD:=MyRs("IDPROD"))
            TbEstP.IDLOJA = mvarIDLOJA
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
            
'            If TbEstP.SLDATUAL < 0 Or TbEstP.SLDDISPONIVEL < 0 Then
'               If ExibirPergunta("Saldo de produto negativo." & vbNewLine & vbNewLine & "Deseja contiuar?", , pDefaultYes:=False) = vbYes Then
'                  Sql = "Select Sum(QTDITEM) [QTDITEM] "
'                  Sql = Sql & " From SMOVEST"
'                  Sql = Sql & " Where IDLOJA = " & mvarIDLOJA
'                  Sql = Sql & " And IDEST= " & TbEstP.IDEST
'                  Sql = Sql & " And IDPROD= " & TbEstP.IDPROD
'                  If mvarSys.xDb.AbreTabela(Sql, MyRsAj) Then
'                     TbEstP.SLDATUAL = xVal(MyRsAj("QTDITEM") & "")
'                     TbEstP.SLDDISPONIVEL = xVal(MyRsAj("QTDITEM") & "")
'                     Queries.Add TbEstP.QrySave
'                  End If
'               Else
'                  GoTo Saida
'               End If
'            End If
                                       
            Set TbMovE = CriarObjeto("BANCO_3R.TB_SMOVEST")
            With TbMovE
               Set .xDb = mvarSys.xDb
               Sql = "TPDOC='CV'"
               Sql = Sql & " And IDVENDA=" & MyRs("IDVENDA") & ""
               Sql = Sql & " And ITEMDOC=" & MyRs("IDITEM") & ""
               Sql = Sql & " And IDFOR=1"
               Sql = Sql & " And IDPROD=" & MyRs("IDPROD") & ""
               If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=Sql) Then
                  TbEstP.SLDATUAL = TbEstP.SLDATUAL + xVal(MyRs("QTDVENDA"))
                  TbEstP.SLDDISPONIVEL = TbEstP.SLDDISPONIVEL + xVal(MyRs("QTDVENDA"))
               End If
               Queries.Add TbEstP.QrySave
               
               If TbMovE.Rs.RecordCount > 1 Then
                  Sql = "Select -1* Sum(QTDITEM)as [QTD]"
                  Sql = Sql & " From SMOVEST"
                  Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
                  Sql = Sql & " And IDMOV<>" & .IDMOV
                  Sql = Sql & " And TPDOC='CV'"
                  Sql = Sql & " And IDVENDA=" & MyRs("IDVENDA") & ""
                  Sql = Sql & " And ITEMDOC=" & MyRs("IDITEM") & ""
                  Sql = Sql & " And IDFOR=1"
                  Sql = Sql & " And IDPROD=" & MyRs("IDPROD") & ""

                  If mvarSys.xDb.AbreTabela(Sql, MyRsAux) Then
                     TbEstP.SLDATUAL = TbEstP.SLDATUAL + xVal(MyRsAux("QTD"))
                     TbEstP.SLDDISPONIVEL = TbEstP.SLDDISPONIVEL + xVal(MyRsAux("QTD"))
                     Queries.Add TbEstP.QrySave
                  End If
                  Queries.Add "Delete From SMOVEST Where IDLOJA=" & mvarIDLOJA & "AND IDMOV<>" & .IDMOV & " And " & Sql
               End If
               .IDLOJA = mvarIDLOJA
               .IDEST = nIDEST
               .IDPROD = MyRs("IDPROD") & ""
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
               .OBS = mvarMe.TxtNOME.Text
'XXX              .SALDO = CalcularSaldoEst
'               .SALDO = SqlNum(0, True)
               Queries.Add .QrySave
            End With
            
            MyRs.MoveNext
         Wend
      End If
      
   End If

   If Queries.Count = 0 Then
      FecharVenda = True
   Else
      FecharVenda = mvarSys.xDb.Executa(Queries)
   End If
   
   If FecharVenda Then
      If bExibeFatura Then
         'Call ExibirFatura(0, mvarIDVENDA, mvarIDATENDIMENTO, mvarIDCLIENTE)
         Call CarregaFechamento
      End If
   Else
      Call ExibirInformacao("Erro ao fechar venda.")
   End If
   
Saida:
   Set Queries = Nothing
   Set TbVenda = Nothing
   Set TbConta = Nothing
   Set TbFLAN = Nothing
   Set TbPgto = Nothing
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
               If pIDCONDPGTO = 1 Then pIDCONDPGTO = 2
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
               .DTVENCIMENTO = pDTVENDA 'DateAdd("d", 1, CDate(pDTVENDA))
               .NPARCELA = 1 'pNParcela
               .tParcela = 1 'pIDCONDPGTO - 1
               
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
   Dim nVLTroco   As Currency
   Dim nVLDivida  As Currency
   Dim nVLPgto    As Currency
   Dim Sql        As String
   
   If Not VerificaVenda() Then Exit Function
   
   bExibeFatura = False
   
   nVLDivida = xVal(mvarMe.TxtVLVENDA.Text) - xVal(mvarMe.TxtVLPGTO.Text)
   
   Sql = "Select Sum(VALOR) [VALOR]"
   Sql = Sql & " From FFATURA "
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And IDVENDAORIGEM=" & mvarIDVENDA
   If mvarIDVENDA > 0 Then
      If mvarSys.xDb.AbreTabela(Sql) Then
         If xVal(mvarSys.xDb.RsAux("VALOR") & "") = 0 Then
            nVLDivida = xVal(mvarMe.TxtVLVENDA.Text) - xVal(mvarMe.TxtVLPGTO.Text)
         Else
            nVLDivida = nVLDivida - xVal(mvarSys.xDb.RsAux("VALOR") & "")
         End If
      End If
   End If
   nVLPgto = xVal(mvarMe.TxtVLPGTO.Text)
   nVLTroco = xVal(mvarMe.TxtVLPGTO.Text) - xVal(mvarMe.TxtVLVENDA.Text)
   With mvarMe.GrdPagamento
      If .RowCount <= 1 Then
         If nVLTroco = 0 Then
            If vbYes = ExibirPergunta("O valor de pagamento é zero." & vbNewLine & vbNewLine & "Deseja continuar?") Then
               If Not SenhaMestre(mvarSys) Then
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
            
            If nVLDivida > 0 Then
               If vbYes = ExibirPergunta("O Sistema detectou falta de pagamento de R$ " & ValBr(nVLDivida) & "." & vbNewLine & vbNewLine & "Deseja gerar Fatura?", "Faturas", True) Then
                  bExibeFatura = True
                  VerificaFechamento = True
               Else
                  Call ExibirInformacao("Pagamento inválido.")
                  If .Visible And .Enabled Then
                     .SetFocus
                  End If
                  Exit Function
               End If
            ElseIf nVLDivida < 0 Then
               Call ExibirInformacao("Pagamento inválido.")
               .SetFocus
               Exit Function
            End If
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
      If Not bExibeFatura Then
         If nVLTroco < 0 Then
            If mvarMe.GrdAtendimento.RowCount > 1 Then
               If mvarMe.GrdAtendimento.CellValue(1, "TIPO") = "Fat." Then
                  Call ExibirInformacao("Pagamento inválido." & vbNewLine & "Pagamento de Fatura deve ser total")
                  bExibeFatura = False
                  Exit Function
               End If
            End If
            If nVLDivida <> 0 Then
               If vbYes = ExibirPergunta("O Sistema detectou falta de pagamento de R$ " & ValBr(nVLDivida) & "." & vbNewLine & vbNewLine & "Deseja gerar Fatura?", "Faturas", True) Then
                  bExibeFatura = True
               Else
                  Call ExibirInformacao("Pagamento inválido.")
                  If mvarMe.GrdPagamento.RowCount > 1 Then
                     .SetFocus
                     Call .SetCurCell(.RowCount - 1, "VLPGTO")
                  Else
                     mvarMe.TxtVLPGTO.SetFocus
                  End If
                  Exit Function
               End If
            End If
         End If
      End If
   End With
   
   VerificaFechamento = True
End Function
Private Sub ExibirAtendimento(pIDATEND As Integer)
   Dim TlAtend As Object
   
   If pIDATEND = 0 Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   Set TlAtend = CriarObjeto("Calendario3R.TL_Atendimento")
   With TlAtend
      Set .Sys = mvarSys
      'Set .CalControl = mvarMe.CalendarControl
      'Set .CalEvent = ModEvent
      .IDLOJA = mvarIDLOJA
      '.IDEVENTO = ModEvent.Id
      '.IDSALA = ModEvent.ScheduleID
      .IDATENDIMENTO = pIDATEND
      .Show vbModal
   End With
   Set TlAtend = Nothing
   Screen.MousePointer = vbDefault
End Sub
Private Sub ExibirItensAtend(lRow As Long)
   Dim MyItens As TL_ITENSATEND
   
   If mvarMe.GrdAtendimento.RowCount <= 1 Then Exit Sub
   If mvarMe.GrdAtendimento.CellValue(lRow, "TIPO") <> "Atend." Then Exit Sub
   
   Set MyItens = New TL_ITENSATEND
   With MyItens
      Set .Sys = mvarSys
      .IDLOJA = mvarIDLOJA
      .IDATEND = xVal(mvarMe.GrdAtendimento.CellValue(lRow, "IDATENDIMENTO"))
      .DeFaultIDVENDA = xVal(mvarMe.TxtIDVENDA.Text)
      
      .Show
      
      mvarMe.GrdAtendimento.Redraw = False
      Call PopulaGrdAtendimento(pIDATEND:=.IDATEND, bRefreshRow:=True)
      On Error Resume Next
      'Call mvarMe.GrdAtendimento.EnsureVisibleCell(lRow, "TIPO")
      Call mvarMe.GrdAtendimento.SetCurCell(lRow, "TIPO")
      mvarMe.GrdAtendimento.Redraw = True
   End With
Exit Sub
   
'   Set MyItens = Nothing
'
'   Dim MYADD   As Object
'   Dim Sql     As String
'   Dim MySel   As Collection
'   Dim MyPre   As Collection
'   Dim MyRs    As Object
'   Dim i       As Integer
'   Dim nIDATEND As Integer
'   Dim n       As Object
'
'   If lRow <= 0 Then Exit Sub
'
'   nIDATEND = xval(mvarMe.GrdAtendimento.CellValue(lRow, "IDATENDIMENTO"))
'   Sql = "Select IDLOJA, IDATENDIMENTO, IDSESSAO, IDTPSERVICO, IDTPTRATAMENTO, IDAREA, IDVENDA"
'   Sql = "Select IDSESSAO, IDTPSERVICO, IDTPTRATAMENTO, IDAREA, IDVENDA"
'   Sql = Sql & " From OSESSAO"
'   Sql = Sql & " Where IDLOJA = " & mvarIDLOJA
'   Sql = Sql & " And IDATENDIMENTO= " & nIDATEND
'
'   If mvarsys.xdb.AbreTabela(Sql, MyRs) Then
'      Set MyPre = New Collection
'      While Not MyRs.EOF
'         If Trim(MyRs("IDVENDA") & "") <> "" Then
'            If Not ExisteItem(MySel, "k" & (MyRs("IDSESSAO") & "")) Then
'               MyPre.Add MyRs("IDSESSAO") & "", "k" & MyRs("IDSESSAO") & ""
'            End If
'         End If
'         MyRs.MoveNext
'      Wend
'      Set MYADD = CriarObjeto("XActive.DsAdd")
'      With MYADD
'         .xdb = mvarsys.xdb
'         .Caption = "Itens do Atendimento."
'         '.Chaves.Add "IDLOJA", "IDLOJA"
'         '.Chaves.Add "IDATENDIMENTO", "IDATENDIMENTO"
'         .Chaves.Add "IDSESSAO", "IDSESSAO"
'         '.Chaves.Add "IDTPSERVICO", "IDTPSERVICO"
'         '.Chaves.Add "IDTPTRATAMENTO", "IDTPTRATAMENTO"
'         '.Chaves.Add "IDAREA", "IDAREA"
'
'         With .TITULO
'            Call .Add("IDSESSAO", "Item", 8, 0) 'eAlin.Esquerda)
'            Call .Add("IDTPSERVICO", "Serviço", 10, 2) 'eAlin.Centro)
'            Call .Add("IDTPTRATAMENTO", "Tratamento", 10, 2) 'eAlin.Centro)
'            Call .Add("IDAREA", "Área", 20, 0) 'eAlin.Esquerda)
'         End With
'         .CheckBox = True
'         .QUERY = Sql
'
'         Set .PreSelecao = MyPre
'
'         .Show
'         Set MySel = .Selecao
'
'
'         Sql = "Update OSESSAO "
'         Sql = Sql & " Set IDVENDA= Null"
'         Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
'         Sql = Sql & " And IDATENDIMENTO=" & SqlNum(nIDATEND)
'         Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
'         Call mvarsys.xdb.Executa(Sql)
'         If MySel.Count > 0 Then
'            For Each n In MySel
'               Sql = "Update OSESSAO "
'               Sql = Sql & " Set IDVENDA= " & mvarIDVENDA
'               Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
'               Sql = Sql & " And IDATENDIMENTO=" & SqlNum(nIDATEND)
'               Sql = Sql & " And IDSESSAO=" & SqlNum(n("IDSESSAO"))
'               Call mvarsys.xdb.Executa(Sql)
'            Next
'         End If
'      End With
'   End If
End Sub
Private Sub mvarMe_CmdIDCLIENTEClick()
   Call ShowCliente(mvarSys, mvarIDLOJA, mvarIDCLIENTE)
End Sub
Private Sub mvarMe_CmdLovCliClick()
   Dim MyCliente As Object
   
   Set MyCliente = CriarObjeto("Contato3R.TL_NovoLov", False)
   With MyCliente
      Set .Sys = mvarSys
      '.Filtro = Trim(mvarMe.TxtNOME.Text)
      Call .ShowClientes
      If .IdSelecionado <> 0 Then
         Call PopulaCliente(.IdSelecionado)
      End If
   End With
   Set MyCliente = Nothing
End Sub

Private Sub mvarMe_CmdOkClick()
   Dim bAux As Boolean
   bAux = mvarMe.CmdFechar.Enabled
   mvarMe.CmdFechar.Enabled = False
   DoEvents
   If Not VerificaVenda() Then Exit Sub
   'If Not bVendaFechada And xVal(mvarIDVENDA) <> 0 Then
      If FecharVenda Then
         Unload mvarMe
      Else
         mvarMe.CmdFechar.Enabled = bAux
      End If
   'Else
   '   If SalvarVenda Then
   '      Unload mvarMe
   '   Else
   '      mvarMe.CmdFechar.Enabled = bAux
   '   End If
   'End If
End Sub
Private Function VerificaVenda() As Boolean
   Dim i As Long
   Dim bDin As Boolean
   Dim bFatura As Boolean
   
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
   If mvarMe.CmbIDFUNCIONARIO.ListIndex < 0 Then
      Call ExibirInformacao("Vendedor inválido." & vbNewLine & "Favor verificar.")
      mvarMe.CmbIDFUNCIONARIO.SetFocus
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
   
   bFatura = False
   For i = 1 To mvarMe.GrdAtendimento.RowCount - 1
      bFatura = (mvarMe.GrdAtendimento.CellValue(i, "TIPO") = "Fat.")
      If bFatura Then Exit For
   Next
   
   With mvarMe.GrdVenda
      If Not bFatura Then
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
      End If
   End With
   
   If xVal(mvarMe.TxtVLVENDA) < xVal(mvarMe.TxtVLPGTO) Then
      Call ExibirInformacao("Valor de Pagamento não confere com o Total da Venda." & vbNewLine & vbNewLine & "Diferença: R$ " & ValBr(xVal(mvarMe.TxtVLPGTO) - xVal(mvarMe.TxtVLVENDA)))
      Exit Function
   End If
   
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
   Dim sIN        As String
   Dim TbVenda    As Object 'TB_CVENDA
   Dim TbAtend    As Object 'TB_OATENDIMENTO_VENDA
   Dim TbItens    As Object 'TB_CITENSVENDA
   Dim TbPgto     As Object 'TB_CPGTOSVENDA
   Dim TbFLAN     As Object 'TB_FLAN
   Dim bServico   As Boolean
   Dim bNFE       As Boolean
   Dim SqlAtend   As String
   Dim MyRs       As Object
   Dim QtdServ1   As Integer
   Dim cItensPago As Collection
   Dim nIDPROD    As Double
                                 
   '* Cliente
   Call SalvarCliente
   
   '* Atendimento
   Call PopulaClasse_Venda(TbVenda)
   bResult = True
   If TbVenda.isDirt Then
      bResult = TbVenda.Salvar
   End If
   bServico = False
   If bResult Then
      mvarIDVENDA = TbVenda.IDVENDA
      mvarMe.TxtIDVENDA.Text = StrZero(TbVenda.IDVENDA, mvarSys.Propriedades("LENIDVENDA"))
      
      Set cQueries = New Collection
      
      '*Atendimentos
      If GetTag(mvarMe.GrdAtendimento, "ISDIRT", 0) = 1 Or GetTag(mvarMe.GrdVenda, "ISDIRT", 0) Then
            
         Set TbAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_VENDA", False)
         Call TbAtend.Limpar
         Set TbAtend.xDb = mvarSys.xDb
         TbAtend.IDLOJA = mvarIDLOJA
         TbAtend.IDVENDA = mvarIDVENDA
         
         Sql = "Select * "
         Sql = Sql & " From FFATURA"
         Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
         Sql = Sql & " And IDVENDA=" & mvarIDVENDA
         If mvarSys.xDb.ExisteReg(Sql) Then
            Sql = "Update FFATURA"
            Sql = Sql & " Set SITFATURA=0"
            Sql = Sql & " , IDVENDA = Null"
            Sql = Sql & " , ALTERSTAMP=1"
            Sql = Sql & " , TIMESTAMP=GetDate()"
            Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
            Sql = Sql & " And IDVENDA=" & mvarIDVENDA
            cQueries.Add Sql
         End If
         SqlAtend = ""
         With mvarMe.GrdAtendimento
            For i = 1 To .RowCount - 1
               If .CellValue(i, "TIPO") = "Atend." Then
                  If Val(.CellValue(i, "IDATENDIMENTO")) <> 0 Then
                     SqlAtend = SqlAtend & IIf(SqlAtend = "", "", ", ") & xVal(.CellText(i, "IDATENDIMENTO"))
                  End If
               End If
            Next
            SqlAtend = IIf(Trim(SqlAtend) = "", "-1", SqlAtend)
            
            Sql = "Select * "
            Sql = Sql & " From OATENDIMENTO"
            Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
            Sql = Sql & " And IDVENDA=" & mvarIDVENDA
            Sql = Sql & " And Not IDATENDIMENTO In (" & SqlAtend & ")"
            If mvarSys.xDb.ExisteReg(Sql) Then
               cQueries.Add TbAtend.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA, Ch_Where:=" Not IDATENDIMENTO In (" & SqlAtend & ")")
            Else
               cQueries.Add TbAtend.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA)
            End If
            For i = 1 To .RowCount - 1
               If Val(.CellValue(i, "IDATENDIMENTO")) <> 0 Then
                  If .CellValue(i, "TIPO") = "Atend." Then
                     TbAtend.IDLOJA = mvarIDLOJA
                     TbAtend.IDVENDA = mvarIDVENDA
                     TbAtend.IDATENDIMENTO = .CellText(i, "IDATENDIMENTO")
                     cQueries.Add TbAtend.QrySave
                  ElseIf .CellValue(i, "TIPO") = "Fat." Then
                     Sql = "Update FFATURA"
                     Sql = Sql & " Set SITFATURA=10"
                     Sql = Sql & " , IDVENDA=" & mvarIDVENDA
                     Sql = Sql & " , ALTERSTAMP=1"
                     Sql = Sql & " , TIMESTAMP=GetDate()"
                     Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
                     Sql = Sql & " And IDFATURA=" & .CellText(i, "IDATENDIMENTO")
                     cQueries.Add Sql
                  End If
               End If
            Next
         End With
         'cQueries.Add TBAtend.QrySave
      End If
      
      '* Itens
      QtdServ1 = 0
      If GetTag(mvarMe.GrdVenda, "ISDIRT", 0) = 1 Then
         Set TbItens = CriarObjeto("BANCO_3R.TB_CITENSVENDA", False)
         TbItens.Limpar
         Set TbItens.xDb = mvarSys.xDb
         TbItens.IDLOJA = mvarIDLOJA
         TbItens.IDVENDA = mvarIDVENDA
         
         With mvarMe.GrdVenda
            cQueries.Add TbItens.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA, Ch_Where:=" IDITEM>" & .RowCount - 1)
            For i = 1 To .RowCount - 1
               If xVal(.CellValue(i, "Total")) <> 0 Then
                  If (Val(.CellValue(i, "ESERVICO")) = 1) Then
                     QtdServ1 = QtdServ1 + xVal(.CellValue(i, "QTDVENDA"))
                  End If
                  bServico = bServico Or (Val(.CellValue(i, "ESERVICO")) = 1)
                  TbItens.IDLOJA = mvarIDLOJA
                  TbItens.IDVENDA = mvarIDVENDA
                  TbItens.IDITEM = i
                  
                  TbItens.IDPROD = Val(.CellValue(i, "IDPROD"))
                  TbItens.NMPROD = .CellText(i, "IDPROD")
                  TbItens.QTDVENDA = xVal(.CellValue(i, "QTDVENDA"))
                  TbItens.UNIDCONTROLE = .CellValue(i, "UNIDCONTROLE")
                  TbItens.VLUNIT = xVal(.CellValue(i, "VLUNIT"))
                  TbItens.IDPACOTE = xVal(.CellValue(i, "IDPACOTE"))
                  TbItens.VLDESC = xVal(.CellValue(i, "VLDESC"))
                  TbItens.DSCDESC = .CellValue(i, "DSCDESC")
                                    
                  cQueries.Add TbItens.QrySave
               End If
            Next
         End With
      Else
         With mvarMe.GrdVenda
            For i = 1 To .RowCount - 1
               If (Val(.CellValue(i, "ESERVICO")) = 1) Then
                  QtdServ1 = QtdServ1 + xVal(.CellValue(i, "QTDVENDA"))
               End If
               If xVal(.CellValue(i, "Total")) <> 0 Then
                  bServico = bServico Or (Val(.CellValue(i, "ESERVICO")) = 1)
               End If
               If bServico Then Exit For
            Next
         End With
      End If
      
      Sql = "Select * "
      Sql = Sql & " From CPACOTE"
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And IDVENDA=" & mvarIDVENDA
      Sql = Sql & " And IDPACOTE>" & mvarPacotes.Count
      If mvarSys.xDb.ExisteReg(Sql) Then
         Dim TbPac  As Object
         Dim vPac   As Variant
         
         Set TbPac = CriarObjeto("BANCO_3R.TB_CPACOTE", False)
         Call TbPac.Limpar
         Set TbPac.xDb = mvarSys.xDb
         cQueries.Add TbPac.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA, Ch_Where:=" IDPACOTE>" & mvarPacotes.Count)
         If mvarPacotes.Count > 0 Then
            For Each vPac In mvarPacotes
               vPac.IDLOJA = mvarIDLOJA
               vPac.IDVENDA = mvarIDVENDA
               vPac.IDCLIENTE = mvarIDCLIENTE
               vPac.DTEMISSAO = Format(mvarMe.CmbDTVENDA.Value, "dd/mm/yyyy")
               cQueries.Add vPac.QrySave
            Next
         End If
      End If
      
      If Trim(SqlAtend) <> "" And QtdServ1 > 0 Then
         Sql = "Select S.IDLOJA, S.IDATENDIMENTO, S.IDSESSAO, S.IDVENDA" & vbNewLine
         Sql = Sql & ", isNull(IDTPSERVICO,0) [IDTPSERVICO], isNull(IDTPTRATAMENTO,0) [IDTPTRATAMENTO], isNull(IDAREA,0) [IDAREA]" & vbNewLine
         Sql = Sql & " From OATENDIMENTO A " & vbNewLine
         Sql = Sql & " Join OSESSAO S On A.IDLOJA=S.IDLOJA AND A.IDATENDIMENTO=S.IDATENDIMENTO And S.IDTPSERVICO<>1" & vbNewLine
         Sql = Sql & " And (isNull(S.IDVENDA,0)=0 Or isNull(S.IDITEM,0)=0)"
         Sql = Sql & " Where A.IDLOJA= " & mvarIDLOJA & vbNewLine
         Sql = Sql & " And A.IDATENDIMENTO In (" & SqlAtend & ")" & vbNewLine
         If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
            Set cItensPago = New Collection
            If MyRs.RecordCount > QtdServ1 Then
               MyRs.MoveFirst
               
               While Not MyRs.EOF
                  '*******
                  '* Verificar se os Itens não cobrados nesta venda já se encontram pagos
                  Sql = "Select * "
                  Sql = Sql & " From OATENDIMENTO_VENDA"
                  Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
                  Sql = Sql & " And IDATENDIMENTO=" & MyRs("IDATENDIMENTO")
                  Sql = Sql & " And IDVENDA=" & MyRs("IDVENDA")
                  If mvarSys.xDb.ExisteReg(Sql) Then
                     QtdServ1 = QtdServ1 + 1
                     If Not ExisteItem(cItensPago, CStr(MyRs("IDSESSAO"))) Then
                        cItensPago.Add MyRs("IDSESSAO"), CStr(MyRs("IDSESSAO"))
                     End If
                  End If
                  MyRs.MoveNext
               Wend
            End If
            MyRs.MoveFirst
            If MyRs.RecordCount <= QtdServ1 Then
               While Not MyRs.EOF
                  If Not ExisteItem(cItensPago, CStr(MyRs("IDSESSAO"))) Then
                     nIDPROD = GetProdutoVenda(MyRs("IDTPSERVICO"), MyRs("IDTPTRATAMENTO"), MyRs("IDAREA"))
                     If nIDPROD > 0 Then
                        Sql = "Update OSESSAO "
                        Sql = Sql & " Set IDVENDA = " & mvarIDVENDA
                        Sql = Sql & " , IDITEM = " & CItensNew(CStr(nIDPROD)).IDITEM
                        Sql = Sql & " , ALTERSTAMP=1"
                        Sql = Sql & " , TIMESTAMP=GetDate()"
                        Sql = Sql & " Where IDLOJA= " & MyRs("IDLOJA")
                        Sql = Sql & " And IDATENDIMENTO=" & MyRs("IDATENDIMENTO")
                        Sql = Sql & " And IDSESSAO=" & MyRs("IDSESSAO")
                        'Sql = Sql & " And (IDVENDA is Null Or IDVENDA=0)"
                        cQueries.Add Sql
                     End If
                  End If
                  MyRs.MoveNext
               Wend
            Else
              Sql = "É necessário identificar itens desta venda"
              Call ExibirInformacao(Sql, "Vendas")
              Call ExibirItensAtend(mvarMe.GrdAtendimento.CurRow)
            End If
         End If
      End If
      
      '* Pagamentos
      If GetTag(mvarMe.GrdPagamento, "ISDIRT", 0) = 1 Then
         With mvarMe.GrdPagamento
            For i = 1 To .RowCount - 1
               Set TbPgto = CriarObjeto("BANCO_3R.TB_CPGTOSVENDA", False)
               Call TbPgto.Limpar
               Set TbPgto.xDb = mvarSys.xDb
               bAchou = TbPgto.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA, Ch_IDPGTO:=.CellValue(i, "IDPGTO"))
            
               If xVal(.CellValue(i, "VLPGTO")) <> 0 Then
                  sIN = IIf(Trim(sIN) = "", "", ", ") + sIN
                  
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
                        Sql = Sql & " IDVENDA = " & SqlNum(mvarIDVENDA)
                        Sql = Sql & " And IDPGTO = " & SqlNum(.CellValue(i, "IDPGTO"))
                        Sql = Sql & " And NPARCELA = " & SqlNum(iParcela)
                        
                        Set TbFLAN = CriarObjeto("BANCO_3R.TB_FLAN", False)
                        With TbFLAN
                           .Limpar
                           Set .xDb = mvarSys.xDb
                           If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=Sql) Then
                              'If ExibirPergunta("Existe lançamento associado ao pagamento." & vbNewLine & vbNewLine & "Deseja atualizar lançamento financeiro?") Then
                                 Dim TbConta As Object 'TB_FCCORRENTE
                                 If DefineContaLAN(TbConta, TbPgto.IDFORMAPGTO, mvarSys) Then
                                    If TbConta.IDCONTA <> .IDCONTA Then
                                       .FLGDELETE = 1
                                       cQueries.Add .QrySave
                                       
                                       'If .AlterarChave(TbFLAN.IDLOJA, TbConta.IDCONTA, TbFLAN.IDLAN) Then
                                        .IDCONTA = TbConta.IDCONTA
                                       'End If
                                    End If
                                 End If
                                 .Valor = TbPgto.VLPGTO - IIf(TbPgto.IDFORMAPGTO = 1, TbVenda.VLTROCO, 0)
                                 Call DefineParcela(TbFLAN, TbVenda.DTVENDA, TbPgto.IDFORMAPGTO, TbPgto.IDCONDPGTO, TbPgto.VLPGTO, iParcela)
                                 .FLGDELETE = 0
                                 cQueries.Add .QrySave
                              'End If
                           End If
                        End With
                     Next
                  Else
                     cQueries.Add TbPgto.QryInsert
                  End If
               End If
            Next
            
            If Trim(sIN) <> "" Or .RowCount <= 1 Then
               If Trim(sIN) = "" Then sIN = "0"
               
               Sql = "Select * "
               Sql = Sql & " From FLAN"
               Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
               Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
               Sql = Sql & " And Not IDPGTO In (" & sIN & ")"
               If mvarSys.xDb.ExisteReg(Sql) Then
                  Sql = "Update FLAN"
                  Sql = Sql & " Set IDVENDA=Null"
                  Sql = Sql & " , IDPGTO =Null"
                  Sql = Sql & " , ALTERSTAMP=1"
                  Sql = Sql & " , TIMESTAMP=GetDate()"
                  Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
                  Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
                  Sql = Sql & " And Not IDPGTO In (" & sIN & ")"
                  cQueries.Add Sql
               End If
               
               Sql = "Select * "
               Sql = Sql & " From CPGTOSVENDA"
               Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
               Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
               Sql = Sql & " And Not IDPGTO In (" & sIN & ")"
               If mvarSys.xDb.ExisteReg(Sql) Then
                  Set TbPgto = CriarObjeto("BANCO_3R.TB_CPGTOSVENDA", False)
                  Call TbPgto.Limpar
                  Set TbPgto.xDb = mvarSys.xDb
                  Sql = " Not IDPGTO In (" & sIN & ")"
                  cQueries.Add TbPgto.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA, Ch_Where:=" Not IDPGTO In (" & sIN & ")")
               End If
            Else
               Sql = "Select * "
               Sql = Sql & " From FLAN"
               Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
               Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
               Sql = Sql & " And IDPGTO >=" & SqlNum(i)
               If mvarSys.xDb.ExisteReg(Sql) Then
                  Sql = "Update FLAN"
                  Sql = Sql & " Set IDVENDA=Null"
                  Sql = Sql & " , IDPGTO =Null"
                  Sql = Sql & " , ALTERSTAMP=1"
                  Sql = Sql & " , TIMESTAMP=GetDate()"
                  Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
                  Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
                  Sql = Sql & " And IDPGTO >=" & SqlNum(i)
                  cQueries.Add Sql
               End If
               Sql = "Select * "
               Sql = Sql & " From CPGTOSVENDA"
               Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
               Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
               Sql = Sql & " And IDPGTO >=" & SqlNum(i)
               If mvarSys.xDb.ExisteReg(Sql) Then
                  Set TbPgto = CriarObjeto("BANCO_3R.TB_CPGTOSVENDA", False)
                  Call TbPgto.Limpar
                  Set TbPgto.xDb = mvarSys.xDb
                  Sql = " IDPGTO >=" & SqlNum(i)
                  cQueries.Add TbPgto.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA, Ch_Where:=" IDPGTO >=" & SqlNum(i))
               End If
            End If
         End With
      End If
            
      SalvarVenda = bResult
      If cQueries.Count > 0 Then
         cQueries.Add "set dateformat 'dmy';", before:=1
         If mvarIDFATURA <> 0 Then
            Dim TBFat As Object
            Set TBFat = CriarObjeto("BANCO_3R.TB_FFATURA", False)
            Call TBFat.Limpar
            Set TBFat.xDb = mvarSys.xDb
            If TBFat.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDFATURA:=mvarIDFATURA) Then
               TBFat.SITFATURA = 10
               TBFat.IDVENDA = mvarIDVENDA
               cQueries.Add TBFat.QryUpdate
            End If
         End If
                           
         bNFE = (1 = mvarSys.GetParam(pIDLOJA:=mvarIDLOJA, pCODPARAM:="NFE", Default:="0", pDescricao:="Nota Fiscal Eletrônica", pSECAO:="NFE"))
         If bServico And bNFE Then
            Call SalvarRPS(cQueries, mvarIDCLIENTE)
         End If
         'For i = 1 To cQueries.Count: Debug.Print cQueries(i): Next
         If mvarSys.xDb.Executa(cQueries) Then
            bResult = True
            SalvarVenda = bResult
            Call SalvarEvento
         Else
            Call ExibirInformacao("Erro ao gravar itens e pagamentos.")
         End If
      Else
         'If bServico And mvarSys.GetParam(pCODPARAM:="NFE", Default:="0", pDescricao:="Nota Fiscal Eletrônica", pSECAO:="NFE") = 1 Then  'pCODSIS:="GLOBAL"
         '   Call SalvarRPS(cQueries, mvarIDCLIENTE)
         'End If
      End If
   Else
      Call ExibirInformacao("Erro ao gravar venda.")
   End If
   If bResult Then
      mvarMe.CmdExcluir.Enabled = Not bVendaFechada
      mvarMe.CmdRecibo.Enabled = Not bVendaFechada
   End If
        
Saida:
   Set TbVenda = Nothing
   Set TbItens = Nothing
   Set TbPgto = Nothing
   Set cQueries = Nothing
End Function
Private Sub SalvarRPS(Optional ByRef pQueries As Collection, Optional ByVal pIDCLIENTE As Integer = 0)
   Dim TBRecibo   As Object 'TB_FRECIBO
   Dim TbCli      As Object
   Dim bExiste    As Boolean
   Dim Sql        As String
   Dim i          As Integer
   Dim bNFE       As Boolean
   Dim bNFE_CLI   As Boolean
   Dim bNFE_CPF   As Boolean
   Dim VlTotal    As Currency
   Dim VlAux      As Currency
   Dim VlDin      As Currency
   Dim VlCredDeb  As Currency
   Dim Total      As Currency
   Dim sMesAux    As String
   Dim nUltDia    As Integer
   
   bNFE = (1 = mvarSys.GetParam(pIDLOJA:=mvarIDLOJA, pCODPARAM:="NFE", Default:="0", pDescricao:="Integração NF-e", pSECAO:="FIN"))
   bNFE_CLI = (1 = mvarSys.GetParam(pIDLOJA:=mvarIDLOJA, pCODPARAM:="NFECLI", Default:="0", pDescricao:="Recibo Cliente", pSECAO:="FIN"))
   bNFE_CPF = (1 = mvarSys.GetParam(pIDLOJA:=mvarIDLOJA, pCODPARAM:="NFECPF", Default:="0", pDescricao:="Recibo por CPF", pSECAO:="FIN"))
   VlTotal = 0
   For i = 1 To mvarMe.GrdVenda.RowCount - 1
      If mvarMe.GrdVenda.CellValue(i, "ESERVICO") = 1 Then
         VlTotal = VlTotal + mvarMe.GrdVenda.CellValue(i, "Total")
      End If
   Next
   VlTotal = VlTotal - xVal(mvarMe.TxtVLDESC.Text)
   
   VlCredDeb = 0
   VlDin = 0
   For i = 1 To mvarMe.GrdPagamento.RowCount - 1
      If mvarMe.GrdPagamento.CellValue(i, "IDFORMAPGTO") = 1 Or mvarMe.GrdPagamento.CellValue(i, "IDFORMAPGTO") = 4 Then
         VlDin = VlDin + mvarMe.GrdPagamento.CellValue(i, "VLPGTO")
      Else
         VlCredDeb = VlCredDeb + mvarMe.GrdPagamento.CellValue(i, "VLPGTO")
      End If
   Next
   
   '****************
   '* Se verificar Cliente
   If pIDCLIENTE <> 0 And Not bNFE_CLI Then
      Sql = "Select NFE"
      Sql = Sql & " From OCLIENTE"
      Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
      Sql = Sql & " And IDCLIENTE=" & SqlNum(pIDCLIENTE)
      If mvarSys.xDb.AbreTabela(Sql) Then
         bNFE_CLI = (xVal(mvarSys.xDb.RsAux("NFE") & "") = 1)
         If bNFE_CLI Then bNFE = True
      End If
   End If
   If Not bNFE_CLI And VlCredDeb = 0 And VlDin <> 0 Then
      VlTotal = VlCredDeb
   End If
   
   Sql = "Select *"
   Sql = Sql & " From FRECIBO"
   Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
   Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
   Sql = Sql & " And SERIE=1"
   bExiste = mvarSys.xDb.ExisteReg(Sql)
   
   If VlCredDeb + VlDin > 0 And bNFE Then
      If bExiste Then
         bNFE = (VlCredDeb <> 0)
      Else
         bNFE = (VlTotal > 0)
      End If
   Else
      bNFE_CLI = False
      bNFE = False
   End If
   
   '****************
   '* Se verificar CPF
   If bNFE And bNFE_CPF Then
      Sql = "Select IDCLIENTE, REGISTRO, EMAIL"
      Sql = Sql & " From OCLIENTE"
      Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
      Sql = Sql & " And IDCLIENTE=" & SqlNum(pIDCLIENTE)
      If mvarSys.xDb.AbreTabela(Sql) Then
         bNFE = bNFE And (Trim(UnFormat(mvarSys.xDb.RsAux("REGISTRO") & "")) <> "")
         bNFE = bNFE And (Trim(UnFormat(mvarSys.xDb.RsAux("EMAIL") & "")) <> "")
         If bNFE_CLI And Not bNFE Then bNFE_CLI = False
      End If
   End If
   
   Total = Val(mvarSys.GetParam(pIDLOJA:=mvarIDLOJA, pCODPARAM:="NFEX", Default:="0", pDescricao:="Param NF", pSECAO:="FIN"))
   If bNFE And Not bNFE_CLI And Total > 0 And Not bExiste Then
      '****************
      '* Verificar limite
      Dim dHoje As Date
      Dim sDom As String
      Dim sSab As String
      Dim TotSem As Currency
      Dim TotMes As Currency
            
      Sql = "Select *"
      Sql = Sql & " From CVENDA"
      Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
      Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
      If mvarSys.xDb.AbreTabela(Sql) Then
         dHoje = mvarSys.xDb.RsAux("DTVENDA")
      Else
         dHoje = mvarSys.xDb.SysDate()
      End If
            
      Sql = "Select Sum(VLTOTAL) as [TotMes]"
      Sql = Sql & " From FRECIBO"
      Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
      Sql = Sql & " And SERIE=1"
      Sql = Sql & " And Year(DTEMISSAO)=" & Year(dHoje)
      Sql = Sql & " And Month(DTEMISSAO)= " & Month(dHoje)
      Sql = Sql & " And DTEMISSAO < " & SqlDate(dHoje)
      If mvarSys.xDb.AbreTabela(Sql) Then
         TotMes = Val(mvarSys.xDb.RsAux("TotMes") & "")
      End If
      If TotMes > Total Then
         bNFE = False
      Else
         sDom = Format(DateAdd("d", 1 - Weekday(dHoje), dHoje), "dd/mm/yyyy 00:00")
         sSab = Format(DateAdd("d", 7 - Weekday(dHoje), dHoje), "dd/mm/yyyy 23:59")
         
         Sql = "Select Sum(VLTOTAL) as [TotSem]"
         Sql = Sql & " From FRECIBO"
         Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
         Sql = Sql & " And SERIE=1"
         Sql = Sql & " And DTEMISSAO >= " & SqlDate(sDom)
         If CDate(dHoje) < CDate(sSab) Then sSab = dHoje
         Sql = Sql & " And DTEMISSAO <= " & SqlDate(sSab)
         If mvarSys.xDb.AbreTabela(Sql) Then
            TotSem = Val(mvarSys.xDb.RsAux("TotSem") & "")
         End If
         If TotSem > (Total / 4) Then
            sMesAux = Format(dHoje, "/mm/yyyy")
            Sql = "Select Sum(VLTOTAL) as [TotSem]"
            Sql = Sql & " From FRECIBO"
            Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
            Sql = Sql & " And SERIE=1"
            Sql = Sql & " And DTEMISSAO >= " & SqlDate("01" & sMesAux & " 00:00")
            If CDate(dHoje) < CDate(sSab) Then sSab = dHoje
            Sql = Sql & " And DTEMISSAO <= " & SqlDate(sSab)
            If mvarSys.xDb.AbreTabela(Sql) Then
               TotSem = Val(mvarSys.xDb.RsAux("TotSem") & "")
               sMesAux = Format(DateAdd("m", 1, CDate(dHoje)), "/mm/yyyy")
               nUltDia = Day(DateAdd("d", -1, CDate("01" & sMesAux)))
               If TotSem / Day(dHoje) > (Total / nUltDia) Then
                  bNFE = False
               End If
            Else
               bNFE = False
            End If
         Else
            If (TotSem / 6) > (Total / 24) Then
               bNFE = False
            End If
         End If
      End If
   End If
   
   If bNFE Or bNFE_CLI Then
      Set TBRecibo = CriarObjeto("BANCO_3R.TB_FRECIBO")
      With TBRecibo
         Set .xDb = mvarSys.xDb
         Sql = " IDVENDA=" & SqlNum(mvarIDVENDA)
         Sql = Sql & " And SERIE=1"
         bExiste = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=Sql)
         If xVal(mvarMe.TxtVLVENDA.Text) <= 0 Then
            If bExiste Then
               .Status = 2
               .VlTotal = 0
               pQueries.Add .QryUpdate
            Else
               Exit Sub
            End If
         End If
         
         .IDLOJA = mvarIDLOJA
         .IDVENDA = mvarIDVENDA
         .DTEMISSAO = mvarMe.CmbDTVENDA.Value
         .VlTotal = VlTotal
         .Tipo = 0            '0-RPS 1-RPS-M (Conjugada/Mista)
         .SERIE = 1           '1-Serviço 2-Produto 3-Remessa
         .Status = 1          '1-Normal 2-Cancelado
         .NATUREZAOP = 0
         .ASSINATURA = ""
         .REGESPECIALTRIB = 0 '0– Nenhum; 1–Microempresa Municipal; 2–Estimativa; 3–Sociedade de Profissionais; 4–Cooperativa; 5–Microempreendedor Individual (MEI); 6–Art. 33, inc. II, item 8, Lei nº 691/84.
         .SIMPLES = 1         '0-Não-Optante pelo Simples Nacional; 1-Optante pelo Simples Nacional
         .TRIBUTACAO = 1      '1-Tributação no Município; 2-Tributação Fora do Município; 3-Operação Isenta; 4-Operação Imune; 5-Operação Suspensa por Decisão Judicial; 6-Operação Suspensa por Decisão Administrativa
         .INCENTIVOCULT = 0   '0–Não; 1–Sim.
         .CODSERVFEDERAL = "0602"               '4 dígitos
         .CODSERVMUNIC = StrZero("060201", 20)  '20 dígitos - "72010" "258873"
         .DSCSERV = mvarSys.GetParam(pIDLOJA:=mvarIDLOJA, pCODPARAM:="DSCSERV", Default:="SERVIÇO DE ESTÉTICA", pDescricao:="Desc. Serviço para NF-e", pSECAO:="FIN")
         '.IDRECIBO
         '.IDLOTE
         '.IDRECIBO0
         '.SERIE0
         
         'If .IDCLIENTE = 0 Then .DTCADASTRO = mvarSys.xDb.sysdate
         If .VlTotal <= 0 Then .Status = 2
         If .isDirt Then
            If bExiste Then
               pQueries.Add .QrySave
            ElseIf .VlTotal > 0 Then
               pQueries.Add .QrySave
            End If
         End If
      End With
      Set TBRecibo = Nothing
   Else
      If bExiste Then
         Call ExcluirRPS(mvarIDLOJA, 0, pQueries, mvarIDVENDA)
         
         'Set TBRecibo = CriarObjeto("BANCO_3R.TB_FRECIBO")
         'Set TBRecibo.xDb = mvarSys.xDb
         'Sql = " IDVENDA=" & SqlNum(mvarIDVENDA)
         'Sql = Sql & " And SERIE=1"
         'If TBRecibo.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=Sql) Then
         '   Call ExcluirRPS(mvarIDLOJA, TBRecibo.IDRECIBO, pQueries)
         'End If
      End If
   End If
End Sub
Private Sub SalvarCliente()
   Dim TBCliente As Object 'TB_OCLIENTE
   Dim bClienteOK As Boolean
   
   If Not (mvarMe.TxtNOME.Enabled And mvarMe.TxtTEL1.Enabled) Then Exit Sub
   
   Set TBCliente = CriarObjeto("BANCO_3R.TB_OCLIENTE", False)
   With TBCliente
      Set .xDb = mvarSys.xDb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCLIENTE:=mvarIDCLIENTE) Then
         .IDCOLIGADA = mvarSys.Propriedades("IDCOLIGADA")
         .NOME = Trim(mvarMe.TxtNOME.Text)
         If .TEL2 <> Trim(mvarMe.TxtTEL1.Text) Then
            .TEL1 = Trim(mvarMe.TxtTEL1.Text)
         End If
         If .IDCLIENTE = 0 Then .DTCADASTRO = mvarSys.xDb.SysDate
         bClienteOK = (.NOME <> "") And (.TEL1 & .TEL2 <> "")
         If .isDirt And bClienteOK Then
            Call .Alterar
            mvarIDCLIENTE = .IDCLIENTE
         End If
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
   Sql = Sql & " Join OATENDIMENTO A        On A.IDLOJA=AV.IDLOJA And A.IDATENDIMENTO=AV.IDATENDIMENTO"
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
            pEvent.CustomProperties("EXIBEMSG") = bVendaFechada
            
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
   
   Set pTBVenda = CriarObjeto("BANCO_3R.TB_CVENDA", False)
   
   With pTBVenda
      .xDb = mvarSys.xDb
      
      Call .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA)
      .IDVENDA = mvarIDVENDA
      .IDLOJA = mvarIDLOJA
      If IsDate(.DTVENDA) And IsDate(mvarMe.CmbDTVENDA.Value) Then
         If CDate(.DTVENDA) <> CDate(mvarMe.CmbDTVENDA.Value) Then
            Call SetTag(mvarMe.GrdPagamento, "ISDIRT", 1)
         End If
      End If
      .DTVENDA = mvarMe.CmbDTVENDA.Value
      .IDCLIENTE = mvarIDCLIENTE
      .IDFUNCIONARIO = mvarMe.CmbIDFUNCIONARIO.ItemData(mvarMe.CmbIDFUNCIONARIO.ListIndex)
      .IDMOEDA = IIf(Val(Sys.Propriedades("IDMOEDA")) = 0, 1, Sys.Propriedades("IDMOEDA"))
      If .SITVENDA = "" Then
         .SITVENDA = "00"
      End If
      .VLDESC = xVal(mvarMe.TxtVLDESC.Text)
      .VLPGTO = xVal(mvarMe.TxtVLPGTO.Text)
      If ValorReal(mvarMe.TxtVLTROCO.Text) < 0 Then
         .VLTROCO = 0
      Else
         .VLTROCO = ValorReal(mvarMe.TxtVLTROCO.Text)
      End If
      
      .VLVENDA = xVal(mvarMe.TxtVLVENDA.Text)
      .VLITENS = xVal(mvarMe.TxtVLSUBTOTAL.Text)
      .VLACRESC = xVal(mvarMe.TxtVLACRESC.Text)
      If mvarIDFATVENDA > 0 Then
         .IDFATURA = mvarIDFATVENDA
      End If
      .OBS = mvarMe.TxtOBS.Text
   End With
End Sub

Private Sub mvarMe_CmdPacotesClick()
   Dim TlPacote As TL_VendaPacote
   Set TlPacote = New TL_VendaPacote
   With TlPacote
      Set .Sys = mvarSys
      .IDVENDA = mvarIDVENDA
      .IDCLIENTE = mvarIDCLIENTE
      .DTVENDA = Format(mvarMe.CmbDTVENDA.Value, "dd/mm/yy")
      Set .Pacotes = mvarPacotes
      .Show
      Set mvarPacotes = .Pacotes
      Call VerificarPacotes
   End With
End Sub
Private Sub VerificarPacotes()
   Dim TlPacote As TL_VendaPacote
   Dim vPac As Variant
   Dim sIN  As String
   Dim Sql As String
   Dim i As Integer
   Dim sMsg As String
   
   If mvarPacotes.Count = 0 Then
      For i = mvarMe.GrdVenda.RowCount To 1 Step -1
         If xVal(mvarMe.GrdVenda.CellValue(i, "IDPACOTE")) <> 0 Then
            mvarMe.GrdVenda.RemoveRow i
         End If
      Next
   Else
      If mvarIDVENDA <= 0 Then
         If mvarMe.GrdVenda.RowCount Then
            sIN = ""
            For Each vPac In mvarPacotes
               If InStr(sIN, vPac.IDPROMO) = 0 Then
                  sIN = sIN & IIf(sIN = "", "", ", ") & vPac.IDPROMO
               End If
            Next
            Sql = "Select * " & vbNewLine
            Sql = Sql & " From CPROMO_PROD " & vbNewLine
            Sql = Sql & " Where IDLOJA=" & mvarIDLOJA & vbNewLine
            Sql = Sql & " And IDPROMO in(" & sIN & ")" & vbNewLine
            sIN = ""
            For i = 1 To mvarMe.GrdVenda.RowCount
               If xVal(mvarMe.GrdVenda.CellValue(i, "IDITEM")) > 0 Then
                  If InStr(sIN, mvarMe.GrdVenda.CellValue(i, "IDPROD")) = 0 Then
                     sIN = sIN & IIf(sIN = "", "", ", ") & mvarMe.GrdVenda.CellValue(i, "IDPROD")
                  End If
               End If
            Next
            Sql = Sql & " And IDPROD in(" & sIN & ")"
            If mvarSys.xDb.ExisteReg(Sql) Then
               sMsg = "Esta venda possui itens de serviço." & vbNewLine & vbNewLine
               sMsg = sMsg & "Deseja incluir este Atendimento no Pacote?"
               If vbYes = ExibirPergunta(sMsg, "Venda") Then
                  Call IncluirPacotes(True)
               Else
                  Call IncluirPacotes(False)
               End If
            Else
               Call IncluirPacotes(False)
            End If
         Else
            Call IncluirPacotes(False)
         End If
      Else
         Call IncluirPacotes(False)
      End If
   End If
   For i = 1 To mvarMe.GrdVenda.RowCount - 1
      mvarMe.GrdVenda.CellValue(i, "IDITEM") = i
   Next
End Sub
Private Sub IncluirPacotes(pClear As Boolean)
   Dim TlPacote As TL_VendaPacote
   Dim vPac As Variant
   Dim sIN  As String
   Dim Sql As String
   Dim i As Integer
   
   For i = mvarMe.GrdVenda.RowCount To 1 Step -1
      If xVal(mvarMe.GrdVenda.CellValue(i, "IDPACOTE")) > 0 Then
         mvarMe.GrdVenda.RemoveRow i
      End If
   Next
   If pClear Then
      sIN = ""
      For Each vPac In mvarPacotes
         If InStr(sIN, vPac.IDPROMO) = 0 Then
            sIN = sIN & IIf(sIN = "", "", ", ") & vPac.IDPROMO
         End If
      Next
      Sql = "Select Distinct IDPROD "
      Sql = Sql & " From CPROMO_PROD "
      Sql = Sql & " Where IDLOJA =" & mvarIDLOJA
      Sql = Sql & " And IDPROMO in(" & sIN & ")"
      If mvarSys.xDb.AbreTabela(Sql) Then
         While Not mvarSys.xDb.RsAux.EOF
            For i = mvarMe.GrdVenda.RowCount To 1 Step -1
               If mvarSys.xDb.RsAux("IDPROD") = mvarMe.GrdVenda.CellValue(i, "IDPROD") And xVal(mvarMe.GrdVenda.CellValue(i, "IDPROD")) <> 0 Then
                  mvarMe.GrdVenda.RemoveRow i
               End If
            Next
            mvarSys.xDb.RsAux.MoveNext
         Wend
      End If
   End If
   If Not pClear And mvarIDVENDA = 0 Then
      Call PopulaNewGRDPRODUTO
   End If
   
   Dim nIDPROMO As Integer
   Dim MyRs As Object
   Dim bExiste As Boolean
   nIDPROMO = 0
   mvarMe.TxtVLDESC.Text = "0,00"
   For Each vPac In mvarPacotes
      With mvarMe.GrdVenda
         If nIDPROMO <> vPac.IDPROMO Then
            Sql = "Select Distinct IDPROD, NMPROD, QTDPROD, UNIDCONTROLE, VLUNIT " & vbNewLine
            Sql = Sql & " From CPROMO_PROD " & vbNewLine
            Sql = Sql & " Where IDLOJA= " & mvarIDLOJA & vbNewLine
            Sql = Sql & " And IDPROMO = " & vPac.IDPROMO
            bExiste = mvarSys.xDb.AbreTabela(Sql, MyRs)
            nIDPROMO = vPac.IDPROMO
         End If
         If bExiste Then
            MyRs.MoveFirst
            While Not MyRs.EOF
               Call IncluiGrdLinhaVenda
               i = .RowCount - 1
               .CellValue(i, "IDPACOTE") = vPac.IDPACOTE
               .CellValue(i, "IDITEM") = i
               .CellValue(i, "IDPROD") = MyRs("IDPROD")
               .CellValue(i, "QTDVENDA") = MyRs("QTDPROD")
               .CellValue(i, "UNIDCONTROLE") = MyRs("UNIDCONTROLE")
               .CellValue(i, "VLUNIT") = ValBr(MyRs("VLUNIT"))
               If i = 1 Then
                  .CellValue(i, "VLDESC") = ValBr(vPac.VLDESC)
               Else
                  If .CellValue(i, "IDPACOTE") <> .CellValue(i - 1, "IDPACOTE") Then
                     .CellValue(i, "VLDESC") = ValBr(vPac.VLDESC)
                  End If
               End If
               .CellForeColor(i, "IDPACOTE") = mvarMe.CmdPacotes.ForeColor
               .CellForeColor(i, "IDITEM") = mvarMe.CmdPacotes.ForeColor
               .CellForeColor(i, "IDPROD") = mvarMe.CmdPacotes.ForeColor
               .CellForeColor(i, "QTDVENDA") = mvarMe.CmdPacotes.ForeColor
               .CellForeColor(i, "UNIDCONTROLE") = mvarMe.CmdPacotes.ForeColor
               .CellForeColor(i, "VLUNIT") = mvarMe.CmdPacotes.ForeColor
               .CellForeColor(i, "Total") = mvarMe.CmdPacotes.ForeColor
               .CellForeColor(i, "VLDESC") = mvarMe.CmdPacotes.ForeColor
               .CellForeColor(i, "DSCDESC") = mvarMe.CmdPacotes.ForeColor
               
               '.CellFont(i, "IDPACOTE").Italic = True
               '.CellFont(i, "IDITEM").Italic = True
               '.CellFont(i, "IDPROD").Italic = True
               '.CellFont(i, "QTDVENDA").Italic = True
               '.CellFont(i, "UNIDCONTROLE").Italic = True
               '.CellFont(i, "VLUNIT").Italic = True
               '.CellFont(i, "Total").Italic = True
               '.CellFont(i, "VLDESC").Italic = True
               '.CellFont(i, "DSCDESC").Italic = True
               
               If i = 1 Then
                  mvarMe.TxtVLDESC.Text = ValBr(vPac.VLDESC + xVal(mvarMe.TxtVLDESC.Text))
               Else
                  If .CellValue(i, "IDPACOTE") <> .CellValue(i - 1, "IDPACOTE") Then
                     mvarMe.TxtVLDESC.Text = ValBr(vPac.VLDESC + xVal(mvarMe.TxtVLDESC.Text))
                  End If
               End If
               Call mvarMe_GrdVendaAfterCommitEdit(i, .ColIndex("IDPROD"))
               
               MyRs.MoveNext
            Wend
         End If
      End With
   Next
   For i = 1 To mvarMe.GrdVenda.RowCount - 1
      mvarMe.GrdVenda.CellValue(i, "IDITEM") = i
   Next
End Sub

Private Sub mvarMe_CmdSaldoClick()
   With mvarMe.GrpAtend
      .Top = (mvarMe.Height / 2) - (.Height / 2)
      .Left = (mvarMe.Width / 2) - (.Width / 2)
      .ZOrder 0
      .Visible = True
   End With
End Sub

Private Sub mvarMe_DblClick()
   'Call ExibirMensagemI(mvarsys, "Cliente possui fatura EM ABERTO", mvarMe.PopupControl)
End Sub
Private Sub mvarMe_GrdAtendimentoColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   bDoDefault = False
End Sub

Private Sub mvarMe_GrdAtendimentoCustomDrawCell(ByVal lRow As Long, ByVal lCol As Long, ByVal hdc As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long, ByVal bSelected As Boolean)
   Dim rcTemp As RECT
   Dim bPushed As Boolean

   If lCol = mvarMe.GrdAtendimento.ColCount Then
      ' drawing the ellipsis button
      bPushed = mvarMe.GrdAtendimento.CellValue(lRow, lCol) = "pushed"
      With rcTemp
         .Top = lTop + 1
         .Right = lRight - 1
         .Left = .Right - BTN_SIZE
         .Bottom = .Top + BTN_SIZE
      End With
      DrawFrameControl hdc, rcTemp, DFC_BUTTON, DFCS_BUTTONPUSH Or IIf(bPushed, DFCS_PUSHED, 0)
      InflateRect rcTemp, -3, -2
      If bPushed Then OffsetRect rcTemp, 1, 1
      SetTextColor hdc, TranslateColor(vbWindowText)
      DrawText hdc, "...", -1, rcTemp, DT_VCENTER Or DT_SINGLELINE Or DT_CENTER

      ' drawing the cell text
      With rcTemp
         .Top = lTop + 1
         .Right = lRight - BTN_SIZE - 1
         .Left = lLeft + 1
         .Bottom = lBottom - 1
      End With
      SetTextColor hdc, TranslateColor(IIf(bSelected, vbHighlightText, vbWindowText))
      DrawText hdc, arrCellText(lRow), -1, rcTemp, DT_SINGLELINE Or DT_CENTER
   End If
End Sub

Private Sub mvarMe_GrdAtendimentoDblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   If (lRow = mvarMe.GrdAtendimento.RowCount) Then
      bRequestEdit = False
   Else
      If mvarMe.GrdAtendimento.CellValue(lRow, "TIPO") = "Atend." Then
         If lCol < 4 Then
            Call ExibirAtendimento(xVal(mvarMe.GrdAtendimento.CellValue(lRow, "IDATENDIMENTO")))
         ElseIf lCol = 4 Then
            Call ExibirItensAtend(lRow)
         End If
      ElseIf mvarMe.GrdAtendimento.CellValue(lRow, "TIPO") = "Fat." Then
         Call ExibirFatura(xVal(mvarMe.GrdAtendimento.CellValue(lRow, "IDATENDIMENTO")))
      End If
   
   End If
End Sub
Private Sub ExibirFatura(ByRef pIDFATURA As Long, Optional pIDVENDAORIGEM, Optional pIDATENDIMENTO, Optional pIDCLIENTE As Long, Optional pVLFATURA)
   Dim oNg As NG_Financeiro
   
   On Error GoTo Saida
   
   Set oNg = New NG_Financeiro
   Set oNg.Sys = mvarSys
   Call oNg.ExibirFatura(pIDFATURA, pIDVENDAORIGEM, pIDATENDIMENTO, pIDCLIENTE, pVLFATURA)
   If pIDFATURA <> 0 And mvarIDFATVENDA <> pIDFATURA Then
      mvarIDFATVENDA = pIDFATURA
   End If
Saida:
   mvarMe.CmdCancel.Enabled = True
   mvarMe.CmdOk.Enabled = True
   Screen.MousePointer = vbDefault
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

Private Sub mvarMe_GrdAtendimentoMouseDown(ByVal Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean, ByVal bUnderControl As Boolean)
   Dim i As Integer
   If lCol = mvarMe.GrdAtendimento.ColCount Then
      m_lBtnRow = lRow
      m_lBtnCol = lCol
      If lRow > 0 And lCol > 0 Then ' the user has clicked a cell (not empty space)
         mvarMe.GrdAtendimento.CellBoundary lRow, lCol, m_lBtnLeft, m_lBtnTop, m_lBtnRight, m_lBtnBottom
         For i = 1 To lCol - 1
            If mvarMe.GrdAtendimento.ColVisible(i) Then
               m_lBtnRight = m_lBtnRight + mvarMe.GrdAtendimento.ColWidth(i)
            End If
         Next
         If (x >= m_lBtnRight - BTN_SIZE) And (x < m_lBtnRight - 1) And (y >= m_lBtnTop + 1) And (y < m_lBtnTop + BTN_SIZE) Then
            mvarMe.GrdAtendimento.CellValue(lRow, lCol) = "pushed" ' when you change cell value, iGrid redraw this cell
         End If
      End If
  End If
End Sub

Private Sub mvarMe_GrdAtendimentoMouseEnter(ByVal lRow As Long, ByVal lCol As Long)
   If mvarMe.GrdAtendimento.ColKey(lCol) = "IDATENDIMENTO" Then
      mvarMe.GrdAtendimento.CellFont(lRow, lCol) = m_oFont
      mvarMe.MousePointer = vbCustom
   End If
End Sub

Private Sub mvarMe_GrdAtendimentoMouseLeave(ByVal lRow As Long, ByVal lCol As Long)
   On Error Resume Next
   If mvarMe.GrdAtendimento.ColKey(lCol) = "IDATENDIMENTO" Then
      mvarMe.GrdAtendimento.CellDefaultFont lRow, lCol
      mvarMe.MousePointer = vbDefault
   End If
End Sub

Private Sub mvarMe_GrdAtendimentoMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
   With mvarMe.GrdAtendimento
      If lRow = .RowCount Then
         If bVendaFechada Then
            If vbYes = ExibirPergunta("Venda fechada." & vbNewLine & "Deseja reabrir venda?") Then
               Call mvarMe_CmdChaveClick
            End If
            If bVendaFechada Then
               Exit Sub
            End If
         End If
      
         'If lRow > 1 Then
         '   If Val(.CellValue(lRow - 1, "IDATENDIMENTO")) = 0 Then Exit Sub
         'End If
         Call IncluiGrdLinhaAtendimento(True)
         
         If lRow > 0 Then
            If xVal(.CellValue(lRow, "IDATENDIMENTO")) = 0 Then
               Call .SetCurCell(lRow, "IDATENDIMENTO")
            End If
         End If
      Else
         If mvarMe.GrdAtendimento.ColKey(lCol) = "IDATENDIMENTO" Then
            If lRow > 0 Then
               Call mvarMe_GrdAtendimentoMouseLeave(lRow, lCol)
               If mvarMe.GrdAtendimento.CellValue(lRow, "TIPO") = "Atend." Then
                  Call ExibirAtendimento(xVal(mvarMe.GrdAtendimento.CellValue(lRow, "IDATENDIMENTO")))
               ElseIf mvarMe.GrdAtendimento.CellValue(lRow, "TIPO") = "Fat." Then
                  Call ExibirFatura(mvarMe.GrdAtendimento.CellValue(lRow, "IDATENDIMENTO"))
               End If
            End If
            
         ElseIf mvarMe.GrdAtendimento.ColKey(lCol) = "ITENS" Then
            If m_lBtnRow > 0 And m_lBtnCol > 0 Then ' if the user has clicked a cell
               mvarMe.GrdAtendimento.CellValue(m_lBtnRow, m_lBtnCol) = Empty ' redraw the pushed button
               'For i = 1 To lCol - 1
               '   If mvarMe.GrdAtendimento.ColVisible(i) Then
               '      m_lBtnRight = m_lBtnRight + mvarMe.GrdAtendimento.ColWidth(i)
               '   End If
               'Next
              
               ' checking whether we need to raise ButtonClick:
               If (x >= m_lBtnRight - BTN_SIZE) And (x < m_lBtnRight - 1) And (y >= m_lBtnTop + 1) And (y < m_lBtnTop + BTN_SIZE) Then
                  ' --- here you can place code for the ButtonClick event:
                  ' ! Don't use modal dialog because VB can't process the iGrid1_CustomDrawCell event
                  ' if you will move the modal dialog over your form
                  Call ExibirItensAtend(lRow)
               End If
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
   Dim nAux    As Integer
   
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
            If TbCond.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=Sql, Ch_ORDERBY:="IDCONDPGTO") Then
               nAux = 0
               While Not TbCond.Rs.EOF
                  TbCond.Popula
                  nAux = IIf(nAux < Len(TbCond.CONDPGTO), Len(TbCond.CONDPGTO), nAux)
                  TbCond.Rs.MoveNext
               Wend
               nAux = nAux + 1
               
               .Clear
               TbCond.Rs.MoveFirst
               While Not TbCond.Rs.EOF
                  TbCond.Popula
                  'If (mvarMe.GrdPagamento.CellValue(lRow, "IDFORMAPGTO") <= 2 And TbCond.IDCONDPGTO = 1) Or (mvarMe.GrdPagamento.CellValue(lRow, "IDFORMAPGTO") = 3 And TbCond.IDCONDPGTO <> 1) Or (mvarMe.GrdPagamento.CellValue(lRow, "IDFORMAPGTO") = 4) Or mvarMe.GrdPagamento.RowCount > 2 Then
                     sAux = TbCond.CONDPGTO
                     sAux = sAux & Space(nAux - Len(sAux))
                     .AddItem sAux, CStr(TbCond.IDCONDPGTO)
                  'End If
                  TbCond.Rs.MoveNext
               Wend
            End If
            Set TbCond = Nothing
            
            If .ListCount > 0 Then
               If mvarMe.GrdPagamento.RowCount - 1 >= lRow Then
                  If InArray(Val(mvarMe.GrdPagamento.CellValue(lRow, "IDFORMAPGTO")), Array(1, 2, 4)) Then
                     mvarMe.GrdPagamento.CellValue(lRow, "IDCONDPGTO") = .ItemValue(0)
                  ElseIf Val(mvarMe.GrdPagamento.CellValue(lRow, "IDFORMAPGTO")) = 3 Then
                     mvarMe.GrdPagamento.CellValue(lRow, "IDCONDPGTO") = CStr(2) '.ItemValue(1)
                  End If
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
   Dim nVLPgto As Currency
   Dim nVlAPagar As Currency
   
   With mvarMe
      With .GrdPagamento
         nVLPgto = 0
         For i = 1 To .RowCount
            nVLPgto = nVLPgto + xVal(.CellValue(i, "VLPGTO"))
         Next
      End With
      nVlAPagar = xVal(.TxtVLVENDA.Text)
      nVlAPagar = nVlAPagar - nVLPgto
      .TxtVLPGTO.Text = ValBr(nVLPgto)
      'If mvarIDFATURA <> 0 Then
      '   If Not .TxtVLPGTO.Visible And xVal(.TxtVLTROCO.Text) <> xVal(-1 * nVlAPagar) Then
      '      .TxtVLDESC = ValBr(xVal(.TxtVLVENDA.Text) + xVal(.TxtVLTROCO.Text))
      '   End If
      'End If
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
         If InArray(Val(.CellValue(lRow, "IDFORMAPGTO")), Array(1, 2)) Then
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
         If eResult = igEditResCommit Then
            .CommitEdit
            .CellValue(lRow, lCol) = .Combos.Items(2).ItemValue(vNewValue - 1)
            'Call .DrawCell(lRow, lCol)
            'Call .SetCurCell(lRow, lCol)
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
   Dim Sql As String
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
         If lRow = 1 Then
            Sql = "Select V.IDVENDA, P.IDFORMAPGTO "
            Sql = Sql & " From CPGTOSVENDA P"
            Sql = Sql & " Join CVENDA V On V.IDLOJA=P.IDLOJA And V.IDVENDA=P.IDVENDA"
            Sql = Sql & " Where V.IDLOJA=" & SqlNum(mvarIDLOJA)
            Sql = Sql & " And V.IDCLIENTE=" & SqlNum(mvarIDCLIENTE)
            Sql = Sql & " Order By V.IDVENDA Desc "
            If mvarSys.xDb.AbreTabela(Sql) Then
              If .CellValue(lRow, "IDFORMAPGTO") <> CStr(mvarSys.xDb.RsAux("IDFORMAPGTO")) Then
                  .CellValue(lRow, "IDFORMAPGTO") = CStr(mvarSys.xDb.RsAux("IDFORMAPGTO"))
                  Call mvarMe_GrdPagamentoAfterCommitEdit(lRow, .ColIndex("IDFORMAPGTO"))
                  .CommitEdit
              End If
            End If
         End If
         .Tag = SetTag(.Tag, "ISDIRT", 1)
         
         If .RowCount > 0 And lRow > 0 Then
            If xVal(.CellValue(lRow, "VLPGTO")) = 0 Then
               Call .SetCurCell(lRow, "VLPGTO")
            End If
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
   Dim MyProd  As Object 'TB_SPRODUTO
   
   With mvarMe.GrdVenda
      If .ColKey(lCol) = "IDPROD" Then
         Set MyProd = CriarObjeto("BANCO_3R.TB_SPRODUTO")
         Set MyProd.xDb = mvarSys.xDb
         If MyProd.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPROD:=.CellValue(lRow, "IDPROD")) Then
            .CellValue(lRow, "ESERVICO") = MyProd.ESERVICO
            .CellValue(lRow, "UNIDCONTROLE") = MyProd.UNIDCONTROLE
            .CellValue(lRow, "VLUNIT") = ValBr(MyProd.VLVENDA)
            .CellValue(lRow, "Total") = ValBr(xVal(.CellValue(lRow, "QTDVENDA")) * xVal(.CellValue(lRow, "VLUNIT")))
            Call CalculaVenda
            Call PopulaValores
         End If
         Set MyProd = Nothing
         
      ElseIf .ColKey(lCol) = "QTDVENDA" Or .ColKey(lCol) = "VLDESC" Then
         .CellValue(lRow, "VLDESC") = ValBr(xVal(.CellValue(lRow, "VLDESC")))
         .CellValue(lRow, "Total") = ValBr(xVal(.CellValue(lRow, "QTDVENDA")) * xVal(.CellValue(lRow, "VLUNIT")))
         '.CellValue(lRow, "Total") = ValBr(xVal(.CellValue(lRow, "QTDVENDA")) * xVal(.CellValue(lRow, "VLUNIT")) - xVal(.CellValue(lRow, "VLDESC")))
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
         
         'If .RowCount > .Combos("IDPROD").ListCount Then
         '   Call ExibirAviso("Não existem mais produtos de venda. (" + CStr(.Combos("IDPROD").ListCount) + ")")
         '   Exit Sub
         'End If
         
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
      If xVal(.CellValue(lRow, "IDPACOTE")) <> 0 Then
         bCancel = True
         Call ExibirInformacao("Item pertence a um pacote e não pode ser alterado." & vbNewLine & "Por favor verifique os pacotes da venda.")
         If mvarMe.CmdPacotes.Enabled Then
            mvarMe.CmdPacotes.SetFocus
         End If
         Exit Sub
      End If
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
   Dim i As Integer
   
   If KeyCode = vbKeyDelete Then
      With mvarMe
         If .ActiveControl Is .GrdPagamento _
         Or .ActiveControl Is .GrdVenda _
         Or .ActiveControl Is .GrdAtendimento Then '  And .GrdAtendimento.CurRow > 1) Then
            If bVendaFechada Then
               Call ExibirInformacao("Venda fechada.")
               Exit Sub
            End If
            If mvarMe.ActiveControl Is mvarMe.GrdVenda Then
               lRow = mvarMe.GrdVenda.CurRow
               If xVal(mvarMe.GrdVenda.CellValue(lRow, "IDPACOTE")) <> 0 Then
                  Call ExibirInformacao("Item pertence a um pacote e não pode ser excluído." & vbNewLine & "Por favor verifique os pacotes da venda.")
                  If mvarMe.CmdPacotes.Enabled Then
                     mvarMe.CmdPacotes.SetFocus
                  End If
                  Exit Sub
               End If
            End If
            With .ActiveControl
               lRow = .CurRow
               If lRow <> .RowCount And lRow > 0 Then
                  .RemoveRow (.CurRow)
                  If mvarMe.ActiveControl Is mvarMe.GrdAtendimento Then
                     arrCellText(mvarMe.GrdAtendimento.RowCount) = ""
                     For i = mvarMe.GrdAtendimento.RowCount - 1 To lRow Step -1
                        arrCellText(i) = GetTag(mvarMe.GrdAtendimento.RowTag(i), "ITENS", "0")
                     Next
                  End If
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
     
   ElseIf InArray(KeyCode, Array(vbKeyF5, vbKeyF6, vbKeyF7, vbKeyF8)) Then
      If Not bVendaFechada Then
         If mvarMe.GrdPagamento.CurRow = mvarMe.GrdPagamento.RowCount Then
            Call mvarMe_GrdPagamentoMouseUp(1, 0, 0, 0, mvarMe.GrdPagamento.RowCount, 1, True)
         End If
         If mvarMe.GrdPagamento.RowCount - 1 > mvarMe.GrdPagamento.CurRow Then
            mvarMe.GrdPagamento.CellValue(mvarMe.GrdPagamento.CurRow, "IDFORMAPGTO") = CStr(KeyCode - vbKeyF4)
            Call mvarMe_GrdPagamentoAfterCommitEdit(mvarMe.GrdPagamento.CurRow, mvarMe.GrdPagamento.ColIndex("IDFORMAPGTO"))
         End If
         mvarMe.GrdPagamento.CommitEdit
         mvarMe_GrdPagamentoLostFocus
      End If
   ElseIf KeyCode = vbKeyF9 Then
      If Not bVendaFechada Then
         Call mvarMe_CmdFecharClick
      End If
   ElseIf KeyCode = vbKeyEscape Then
      Call mvarMe_CmdCancelClick
   End If
   
   If Not bVendaFechada Then
      If mvarMe.ActiveControl Is mvarMe.GrdPagamento Then
         If InArray(KeyCode, Array(vbKeyF1, vbKeyF2, vbKeyF3)) Then
            If Trim(UCase(mvarMe.GrdPagamento.CellText(mvarMe.GrdPagamento.CurRow, "IDFORMAPGTO"))) = "CRÉDITO" Then
               mvarMe.GrdPagamento.CellValue(mvarMe.GrdPagamento.CurRow, "IDCONDPGTO") = CStr(KeyCode - 110)
               Call mvarMe_GrdPagamentoAfterCommitEdit(mvarMe.GrdPagamento.CurRow, mvarMe.GrdPagamento.ColIndex("IDCONDPGTO"))
               mvarMe.GrdPagamento.CommitEdit
               mvarMe_GrdPagamentoLostFocus
            End If
         End If
         If mvarMe.GrdPagamento.CurCol = mvarMe.GrdPagamento.ColIndex("IDFORMAPGTO") Then
            If InArray(KeyCode, Array(49, 50, 51, 52)) Then
               mvarMe.GrdPagamento.CellValue(mvarMe.GrdPagamento.CurRow, "IDFORMAPGTO") = CStr(KeyCode - 48)
               Call mvarMe_GrdPagamentoAfterCommitEdit(mvarMe.GrdPagamento.CurRow, mvarMe.GrdPagamento.ColIndex("IDFORMAPGTO"))
               mvarMe.GrdPagamento.CommitEdit
               mvarMe_GrdPagamentoLostFocus
            End If
            If InArray(KeyCode, Array(97, 98, 99, 100)) Then
               mvarMe.GrdPagamento.CellValue(mvarMe.GrdPagamento.CurRow, "IDFORMAPGTO") = CStr(KeyCode - 96)
               Call mvarMe_GrdPagamentoAfterCommitEdit(mvarMe.GrdPagamento.CurRow, mvarMe.GrdPagamento.ColIndex("IDFORMAPGTO"))
               mvarMe.GrdPagamento.CommitEdit
               mvarMe_GrdPagamentoLostFocus
            End If
         
         ElseIf mvarMe.GrdPagamento.CurCol = mvarMe.GrdPagamento.ColIndex("IDCONDPGTO") And mvarMe.GrdPagamento.CurRow > 0 Then
            If Trim(UCase(mvarMe.GrdPagamento.CellText(mvarMe.GrdPagamento.CurRow, "IDFORMAPGTO"))) = "CRÉDITO" Then
               If InArray(KeyCode, Array(49, 50, 51)) Then
                  mvarMe.GrdPagamento.CellValue(mvarMe.GrdPagamento.CurRow, "IDCONDPGTO") = CStr(KeyCode - 47)
                  Call mvarMe_GrdPagamentoAfterCommitEdit(mvarMe.GrdPagamento.CurRow, mvarMe.GrdPagamento.ColIndex("IDCONDPGTO"))
                  mvarMe.GrdPagamento.CommitEdit
                  mvarMe_GrdPagamentoLostFocus
               End If
               If InArray(KeyCode, Array(97, 98, 99)) Then
                  mvarMe.GrdPagamento.CellValue(mvarMe.GrdPagamento.CurRow, "IDCONDPGTO") = CStr(KeyCode - 95)
                  Call mvarMe_GrdPagamentoAfterCommitEdit(mvarMe.GrdPagamento.CurRow, mvarMe.GrdPagamento.ColIndex("IDCONDPGTO"))
                  mvarMe.GrdPagamento.CommitEdit
                  mvarMe_GrdPagamentoLostFocus
               End If
            
            End If
         End If
      End If
   End If
End Sub
Private Sub mvarMe_Activate()
'   Call mvarMe_Load
   
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
   mvarMe.MousePointer = vbDefault
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_Load()
   mvarMe.Icon = mvarSys.MDI.Icon
   '* Monta Tela
   Call SetExistePromo
   Call MontaTela
   Call MontaGrdVenda
   Call MontaGrdPgto
   Call MontaGrdAtendimento
   
   Call CarregaFechamento
   Call PopulaTela
   Call VerificaMensagemI
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub VerificaMensagemI()
   Dim Sql As String
   Dim sMsg As String
   
   Sql = "Select F.IDFATURA, C.NOME"
   Sql = Sql & " From FFATURA F"
   Sql = Sql & " Join OCLIENTE C On C.IDLOJA=F.IDLOJA AND C.IDCLIENTE=F.IDCLIENTE"
   Sql = Sql & " Where F.SITFATURA = 0"
   Sql = Sql & " And F.IDLOJA = " & mvarIDLOJA
   Sql = Sql & " And F.IDCLIENTE = " & mvarIDCLIENTE
   Sql = Sql & " And F.IDATENDIMENTO<>" & mvarIDATENDIMENTO
   
   If mvarSys.xDb.AbreTabela(Sql) Then
      With mvarSys.xDb.RsAux
         If .RecordCount > 0 Then
            sMsg = "Cliente: " & !NOME
            sMsg = Mid(sMsg, 1, 28) & IIf(Len(sMsg) > 30, "...", "") & vbNewLine
            sMsg = sMsg & "Possui fatura " & StrZero(!IDFATURA, 6) & " 'Em Aberto'"
            Call ExibirMensagemI(mvarSys, sMsg)
         End If
      End With
   End If
End Sub
Private Sub CarregaFechamento()
   Dim TbVenda As Object 'TB_CVENDA
   Dim TbAV    As Object 'TB_OATENDIMENTO_VENDA
   
   bVendaFechada = False

   If mvarIDVENDA = 0 And mvarIDATENDIMENTO <> 0 And mvarIDFATURA = 0 Then
      If mvarMe.CmbDTVENDA.Value <> mvarMe.CmbDTVENDA.MinDate Then
         If CDate(Format(mvarMe.CmbDTVENDA.Value, "dd/mm/yyyy")) <> CDate(mvarSys.xDb.SysDate()) Then
            Set TbAV = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_VENDA")
            Set TbAV.xDb = mvarSys.xDb
            If TbAV.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO) Then
               mvarIDVENDA = TbAV.IDVENDA
            End If
         End If
      End If
   End If
   If mvarIDVENDA > 0 Then
      Set TbVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
      Set TbVenda.xDb = mvarSys.xDb
      If TbVenda.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA) Then
         bVendaFechada = (TbVenda.SITVENDA = gSITVENDA_Fechada)
      End If
      Dim TBVenda2 As Object
      'Set TBVenda2 = CriarObjeto("BANCO_3R.TB_CVENDA")
      'Set TBVenda2.xDb = mvarsys.xDb
      'If TBVenda2.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=2884) Then
      '   bVendaFechada = (TBVenda2.SITVENDA = gSITVENDA_Fechada)
      'End If
   
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
      .CmbDTVENDA.Value = .CmbDTVENDA.MinDate  '"01/01/1601"
      
      .TxtVLSUBTOTAL.Font.Bold = False
      .TxtVLSUBTOTAL.Font.Size = 14
      .TxtVLSUBTOTAL.Font.Name = "Courier New"
      .TxtVLSUBTOTAL.Locked = True
      .TxtVLSUBTOTAL.Enabled = False
      
      .TxtVLDESC.Font.Size = 14
      .TxtVLDESC.Font.Name = "Courier New"
      .TxtVLDESC.Locked = True
      .TxtVLDESC.Enabled = False
      
      .TxtVLACRESC.Font.Size = 14
      .TxtVLACRESC.Font.Name = "Courier New"
      .TxtVLACRESC.Locked = False
      .TxtVLACRESC.Enabled = True
      .TxtVLACRESC.Text = "0,00"
      
      
      .TxtVLVENDA.Font.Bold = True
      .TxtVLVENDA.Font.Size = 16
      .TxtVLVENDA.Font.Name = "Courier New"
      .TxtVLVENDA.BackColor = .TxtIDVENDA.BackColor
      .TxtVLVENDA.ForeColor = vbBlue
      .TxtVLVENDA.Enabled = True
      .TxtVLVENDA.Locked = True
      
      
      .TxtVLPGTO.Font.Size = 16
      .TxtVLPGTO.Font.Name = "Courier New"
      .TxtVLPGTO.Enabled = False
      .TxtVLPGTO.Locked = True
      
      .TxtVLTROCO.Font.Size = 16
      .TxtVLTROCO.Font.Name = "Courier New"
      .TxtVLTROCO.Locked = True
            
      .CmdPacotes.Enabled = bExistePromo
      
      .GrdVenda.Enabled = (mvarIDFATURA = 0)
      
      Call MontaCmbFuncionario
   End With
End Sub
Private Sub PopulaTela()
   Dim Sql     As String
   Dim TbAtend As Object 'TB_OATENDIMENTO
   Dim TbVenda As Object 'TB_CVENDA
   Dim TbAV    As Object 'TB_OATENDIMENTO_VENDA
   Dim TBFat   As Object 'TB_FFATURA
   Dim bOk     As Boolean
   Dim sAux    As String
   Dim sIDFat  As String
   
   Call LimparTela
   
   'If mvarIDVENDA = 0 And mvarIDATENDIMENTO <> 0 And mvarIDFATURA = 0 Then
   '   Set TbAV = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_VENDA")
   '   Set TbAV.xDb = mvarsys.xDb
   '
   '   bOk = TbAV.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO)
   '   If bOk Then mvarIDVENDA = TbAV.IDVENDA
   'End If
   
   If mvarIDVENDA > 0 Then
      mvarMe.CmdPacotes.Enabled = True
      mvarMe.CmbDTVENDA.Enabled = False
      If TbVenda Is Nothing Then
         Set TbVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
         Set TbVenda.xDb = mvarSys.xDb
         bOk = TbVenda.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA)
      End If
      If bOk Then
         With TbVenda
            'bVendaFechada = (.SITVENDA = gSITVENDA_Fechada)
                     
            mvarMe.TxtIDVENDA.Text = StrZero(.IDVENDA, mvarSys.Propriedades("LENIDVENDA"))
            mvarMe.CmbDTVENDA.Value = .DTVENDA
                                    
            Call PopulaCliente(.IDCLIENTE)
            Call PopulaGrdAtendimento
            Call PopulaGRDVENDA
            Call PopulaGRDPGTO
            Call PopulaRecibo
            
            Call CalculaVenda

mvarMe.TxtVLDESC.Visible = False
mvarMe.TxtVLACRESC.Visible = False
            
            mvarMe.TxtVLSUBTOTAL.Text = ValBr(.VLITENS)
            mvarMe.TxtVLDESC.Text = ValBr(.VLDESC)
            mvarMe.TxtVLACRESC.Text = ValBr(.VLACRESC)
            mvarMe.TxtVLPGTO.Text = ValBr(.VLPGTO)
            mvarMe.TxtVLVENDA.Text = ValBr(.VLVENDA)
            Call PopulaTroco(.VLTROCO)

mvarMe.TxtVLDESC.Visible = True
mvarMe.TxtVLACRESC.Visible = True
            
            Call LocalizarCombo(mvarMe.CmbIDFUNCIONARIO, .IDFUNCIONARIO, True, True)
            mvarMe.TxtOBS.Text = .OBS

         End With
      End If
      Set TbVenda = Nothing
      
      Dim TlPacote As TL_VendaPacote
      Set TlPacote = New TL_VendaPacote
      With TlPacote
         Set .Sys = mvarSys
         .IDVENDA = mvarIDVENDA
         .IDCLIENTE = mvarIDCLIENTE
         .DTVENDA = Format(mvarMe.CmbDTVENDA.Value, "dd/mm/yy")
         .Popula
         Set mvarPacotes = .Pacotes
      End With
      Set TlPacote = Nothing
   ElseIf mvarIDATENDIMENTO <> 0 Then
      mvarMe.CmdPacotes.Enabled = bExistePromo
      mvarMe.CmbDTVENDA.Enabled = True
      If mvarIDFATURA <> 0 Then
         mvarMe.CmdPacotes.Enabled = False
         'DoEvents
         Set TBFat = CriarObjeto("BANCO_3R.TB_FFATURA")
         Set TBFat.xDb = mvarSys.xDb
         If TBFat.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDFATURA:=mvarIDFATURA) Then
            With TBFat
               Call PopulaCliente(.IDCLIENTE)
               Call PopulaGrdAtendimento(pIDFATURA:=mvarIDFATURA)
               DoEvents
               'mvarMe.TxtVLDESC.Text = ValBr(.VLDESC)
               'mvarMe.TxtVLTROCO.Text = ValBr(-1 * .Valor)
               If CDate(.DTPREV) > CDate(mvarSys.xDb.SysDate()) Then
                  mvarMe.CmbDTVENDA.Value = .DTPREV
               Else
                  mvarMe.CmbDTVENDA.Value = mvarSys.xDb.SysDate()
               End If
            End With
         End If
      Else
         
         Set TbAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
         Set TbAtend.xDb = mvarSys.xDb
         If TbAtend.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO) Then
            With TbAtend
               Call PopulaCliente(.IDCLIENTE)
               Call PopulaGrdAtendimento(pIDATEND:=mvarIDATENDIMENTO, bNewVenda:=True)
            End With
         End If
         Set TbAtend = Nothing
      End If
      Call LocalizarCombo(mvarMe.CmbIDFUNCIONARIO, CStr(mvarIDFUNCIONARIO), True, True)
      'mvarMe.TxtObs.Text ="" '.OBS
      Call CalculaVenda
      Call PopulaValores
      
   Else
      Call PopulaCliente(mvarIDCLIENTE)
      Call CalculaVenda
      Call PopulaValores
   End If
   
   mvarMe.CmdFaturas.Enabled = False
   mvarMe.CmdRecibo.Enabled = False
   If mvarIDVENDA > 0 And bOk Then
      mvarMe.CmdFechar.Enabled = Not bVendaFechada
      mvarMe.CmdOk.Enabled = Not bVendaFechada
      mvarMe.CmdExcluir.Enabled = Not bVendaFechada
      mvarMe.TxtIDCLIENTE.Enabled = Not bVendaFechada
      mvarMe.CmdLovCli.Enabled = Not bVendaFechada
      mvarMe.TxtNOME.Locked = bVendaFechada
      mvarMe.TxtTEL1.Locked = bVendaFechada
      mvarMe.CmdChave.Enabled = bVendaFechada
      mvarMe.CmdPacotes.Enabled = Not bVendaFechada
      
      sAux = ""
      Sql = "Select IDFATURA"
      Sql = Sql & " From FFATURA "
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And IDVENDAORIGEM=" & mvarIDVENDA
      Sql = Sql & " Order By IDFATURA"
      'Sql = Sql & " And IDATENDIMENTO =" & mvarIDATENDIMENTO
      If mvarSys.xDb.AbreTabela(Sql) Then
         mvarMe.CmdFaturas.Enabled = True
         sAux = "Verifique a consistência do pagamento na Fatura: Nº"
         sIDFat = StrZero(mvarSys.xDb.RsAux("IDFATURA"), 6)
      Else
         Sql = "Select IDFATURA "
         Sql = Sql & " From CVENDA "
         Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
         Sql = Sql & " And IDVENDA =" & mvarIDVENDA
         Sql = Sql & " And Year(DTVENDA)>2012"
         Sql = Sql & " And IDFATURA>1"
         If mvarSys.xDb.AbreTabela(Sql) Then
            mvarMe.CmdFaturas.Enabled = True
            sAux = "Verifique a consistência do pagamento na Fatura: Nº"
            sIDFat = StrZero(mvarSys.xDb.RsAux("IDFATURA"), 6)
            
         End If
      End If
      If sAux <> "" Then
         mvarMe.LblFatura.Caption = sAux
         mvarMe.CmdFatura.Caption = sIDFat
         mvarMe.LblFatura.Visible = True
         mvarMe.CmdFatura.Visible = True
      End If
   End If
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
   Dim TbItem As Object 'TB_CITENSVENDA
   Dim sTag As String
   Dim i    As Integer
   Dim Sql  As String

   DoEvents
   Set TbItem = CriarObjeto("BANCO_3R.TB_CITENSVENDA")
   Set TbItem.xDb = mvarSys.xDb
      
   
   If TbItem.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA, Ch_ORDERBY:="IDITEM") Then
      With mvarMe.GrdVenda
         .Clear
         Call IncluiGrdLinhaVenda
         
         While Not TbItem.Rs.EOF
            TbItem.Popula
                  
            Call IncluiGrdLinhaVenda
         
            i = .RowCount - 1
            If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
            
            Call SetTag(sTag, "IDPACOTE", TbItem.IDPACOTE)
            Call SetTag(sTag, "IDITEM", TbItem.IDITEM)
            Call SetTag(sTag, "IDPROD", TbItem.IDPROD)
            Call SetTag(sTag, "QTDVENDA", TbItem.QTDVENDA)
            Call SetTag(sTag, "UNIDCONTROLE", TbItem.UNIDCONTROLE)
            Call SetTag(sTag, "VLUNIT", ValBr(TbItem.VLUNIT))
            Call SetTag(sTag, "Total", ValBr(TbItem.QTDVENDA * TbItem.VLUNIT) - ValBr(TbItem.VLDESC))
            Call SetTag(sTag, "VLDESC", ValBr(TbItem.VLDESC))
            Call SetTag(sTag, "DSCDESC", TbItem.DSCDESC)
            
            .RowTag(i) = sTag
            
            .CellValue(i, "IDPACOTE") = TbItem.IDPACOTE
            .CellValue(i, "IDITEM") = TbItem.IDITEM
            .CellValue(i, "IDPROD") = TbItem.IDPROD
            .CellValue(i, "QTDVENDA") = TbItem.QTDVENDA
            .CellValue(i, "UNIDCONTROLE") = TbItem.UNIDCONTROLE
            .CellValue(i, "VLUNIT") = ValBr(TbItem.VLUNIT)
            .CellValue(i, "Total") = ValBr(TbItem.QTDVENDA * TbItem.VLUNIT - TbItem.VLDESC)
            .CellValue(i, "VLDESC") = ValBr(TbItem.VLDESC)
            .CellValue(i, "DSCDESC") = TbItem.DSCDESC
            If TbItem.IDPACOTE > 0 Then
               .CellForeColor(i, "IDPACOTE") = mvarMe.CmdPacotes.ForeColor
               .CellForeColor(i, "IDITEM") = mvarMe.CmdPacotes.ForeColor
               .CellForeColor(i, "IDPROD") = mvarMe.CmdPacotes.ForeColor
               .CellForeColor(i, "QTDVENDA") = mvarMe.CmdPacotes.ForeColor
               .CellForeColor(i, "UNIDCONTROLE") = mvarMe.CmdPacotes.ForeColor
               .CellForeColor(i, "VLUNIT") = mvarMe.CmdPacotes.ForeColor
               .CellForeColor(i, "Total") = mvarMe.CmdPacotes.ForeColor
               .CellForeColor(i, "VLDESC") = mvarMe.CmdPacotes.ForeColor
               .CellForeColor(i, "DSCDESC") = mvarMe.CmdPacotes.ForeColor
            End If

            Dim MyProd As Object
            Set MyProd = CriarObjeto("BANCO_3R.TB_SPRODUTO")
            Set MyProd.xDb = mvarSys.xDb
            If MyProd.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPROD:=.CellValue(i, "IDPROD")) Then
               .CellValue(i, "ESERVICO") = MyProd.ESERVICO
            End If
            Set MyProd = Nothing

                     
            TbItem.Rs.MoveNext
         Wend
         .RowVisible(1) = True
         '.CellSelected(1, "FOTOTIPO") = True
         .Tag = SetTag(.Tag, "ISDIRT", 0)
      End With
   End If
   Set TbItem = Nothing
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
            If MyTb.IDFORMAPGTO = 3 And MyTb.IDCONDPGTO = 1 Then
               MyTb.IDCONDPGTO = 2
               Call MyTb.Salvar
            End If
            
            Call IncluiGrdLinhaPgto
         
            i = .RowCount - 1
            If i >= 0 Then
               If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
            End If
            Call SetTag(sTag, "IDPGTO", MyTb.IDPGTO)
            Call SetTag(sTag, "IDCONDPGTO", MyTb.IDCONDPGTO)
            Call SetTag(sTag, "IDFORMAPGTO", MyTb.IDFORMAPGTO)
            Call SetTag(sTag, "VLPGTO", ValBr(MyTb.VLPGTO))
            If i >= 0 Then
               .RowTag(i) = sTag
               
               .CellValue(i, "IDPGTO") = MyTb.IDPGTO
               .CellValue(i, "IDFORMAPGTO") = CStr(MyTb.IDFORMAPGTO)
               Call mvarMe_GrdPagamentoAfterCommitEdit(i, .ColIndex("IDFORMAPGTO"))
               .CellValue(i, "IDCONDPGTO") = CStr(MyTb.IDCONDPGTO)
               .CellValue(i, "VLPGTO") = ValBr(MyTb.VLPGTO)
            End If
            MyTb.Rs.MoveNext
         Wend
         If .RowCount > 0 Then
            .RowVisible(1) = True
            '.CellSelected(1, "FOTOTIPO") = True
         End If
         .Tag = SetTag(.Tag, "ISDIRT", 0)
      End With
   End If
   Set MyTb = Nothing
End Sub
Private Sub PopulaGRDPRODUTO(Optional pIDFATURA As Long = 0)
   Dim i As Integer
   Dim QtdServ As Integer
   
   If mvarIDVENDA <> 0 Then
      If bVendaFechada Then
         Exit Sub
      ElseIf CDate(Format(mvarMe.CmbDTVENDA.Value, "dd/mm/yyyy")) < CDate(Format(mvarSys.xDb.SysDate, "dd/mm/yyyy")) Then
         If vbNo = ExibirPergunta("Deseja recalcular venda?", "Vendas") Then
            Exit Sub
         End If
      End If
   End If
   QtdServ = 0
   For i = 1 To mvarMe.GrdAtendimento.RowCount
      QtdServ = QtdServ + xVal(arrCellText(i))
   Next
End Sub
Private Sub PopulaNewGRDPRODUTO(Optional pIDFATURA As Long = 0)
   Dim MyTbServ As Object 'TB_OSESSAO
   Dim MyTbProd As Object 'TB_OATENDIMENTO_PRODUTO
   Dim sTag    As String
   Dim i       As Integer
   Dim Sql     As String
   
   Dim pIDATENDIMENTO As Long
   Dim MyItem  As Object 'TB_CITENSVENDA
   Dim MyProd  As Object 'TB_SPRODUTO
   Dim MyItens As Object 'TB_FITEMFATURA
   Dim nIDPROD As Long
   Dim j       As Integer
   Dim bPago   As Boolean
   Dim sMsg    As String
   
   Set CItensNew = New Collection
   
   If mvarIDVENDA <> 0 Then
      If bVendaFechada Then
         Exit Sub
      ElseIf CDate(Format(mvarMe.CmbDTVENDA.Value, "dd/mm/yyyy")) < CDate(Format(mvarSys.xDb.SysDate, "dd/mm/yyyy")) And mvarIDVENDA > 0 Then
         If vbNo = ExibirPergunta("Deseja recalcular venda?", "Vendas") Then
            Exit Sub
         End If
      End If
   End If
   
   With mvarMe.GrdVenda
      If pIDFATURA = 0 Then
         For j = 1 To mvarMe.GrdAtendimento.RowCount
            If mvarMe.GrdAtendimento.CellValue(j, "TIPO") = "Atend." Then
               pIDATENDIMENTO = Val(mvarMe.GrdAtendimento.CellValue(j, "IDATENDIMENTO"))
                        
               Set MyTbServ = CriarObjeto("BANCO_3R.TB_OSESSAO")
               Set MyTbServ.xDb = mvarSys.xDb
               If MyTbServ.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=pIDATENDIMENTO) Then
                  While Not MyTbServ.Rs.EOF
                     MyTbServ.Popula
                     '***********
                     '* Verifica se item já está pago
                     bPago = False
                     If MyTbServ.IDVENDA <> 0 And MyTbServ.IDITEM <> 0 Then
                        'Sql = "Select * "
                        'Sql = Sql & " From OATENDIMENTO_VENDA"
                        'Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
                        'Sql = Sql & " And IDATENDIMENTO=" & pIDATENDIMENTO
                        'Sql = Sql & " And IDVENDA=" & MyTbServ.IDVENDA
                        'bPago = mvarSys.xDb.ExisteReg(Sql)
                        bPago = True
                     End If
                     If Not bPago Then
                        nIDPROD = GetProdutoVenda(MyTbServ.IDTPSERVICO, MyTbServ.IDTPTRATAMENTO, MyTbServ.IDAREA)
                        If nIDPROD = 0 Then
                           If mvarSys.xDb.AbreTabela("Select DSCTRATAMENTO From OTPTRATAMENTO Where IDLOJA=" & mvarIDLOJA & " And IDTPTRATAMENTO=" & MyTbServ.IDTPTRATAMENTO) Then
                              sMsg = sMsg & ", " & mvarSys.xDb.RsAux("DSCTRATAMENTO")
                           End If
                        Else
                           If Not ExisteItem(CItensNew, CStr(nIDPROD)) Then
                              Set MyItem = CriarObjeto("BANCO_3R.TB_CITENSVENDA")
                              CItensNew.Add MyItem, CStr(nIDPROD)
                              CItensNew(CStr(nIDPROD)).IDITEM = CItensNew.Count
                           End If
                           CItensNew(CStr(nIDPROD)).IDPROD = nIDPROD
                           CItensNew(CStr(nIDPROD)).QTDVENDA = CItensNew(CStr(nIDPROD)).QTDVENDA + 1
                        End If
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
                        If Not ExisteItem(CItensNew, CStr(nIDPROD)) Then
                           Set MyItem = CriarObjeto("BANCO_3R.TB_CITENSVENDA")
                           CItensNew.Add MyItem, CStr(nIDPROD)
                           CItensNew(CStr(nIDPROD)).IDITEM = CItensNew.Count
                        End If
                        CItensNew(CStr(nIDPROD)).IDPROD = nIDPROD
                        CItensNew(CStr(nIDPROD)).QTDVENDA = CItensNew(CStr(nIDPROD)).QTDVENDA + MyTbProd.QTDCOMPRA
                     End If
                     MyTbProd.Rs.MoveNext
                  Wend
               End If
               Set MyTbProd = Nothing
            End If
         Next
         
      Else
         Set MyItens = CriarObjeto("BANCO_3R.TB_FITEMFATURA")
         Set MyItens.xDb = mvarSys.xDb
         If MyItens.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDFATURA:=mvarIDFATURA) Then
            While Not MyItens.Rs.EOF
               MyItens.Popula
               nIDPROD = MyItens.IDPROD
               If nIDPROD > 0 Then
                  If Not ExisteItem(CItensNew, CStr(nIDPROD)) Then
                     Set MyItem = CriarObjeto("BANCO_3R.TB_CITENSVENDA")
                     CItensNew.Add MyItem, CStr(nIDPROD)
                     CItensNew(CStr(nIDPROD)).IDITEM = CItensNew.Count
                  End If
                  CItensNew(CStr(nIDPROD)).IDPROD = nIDPROD
                  CItensNew(CStr(nIDPROD)).QTDVENDA = CItensNew(CStr(nIDPROD)).QTDVENDA + MyItens.QTDITEM
                  CItensNew(CStr(nIDPROD)).UNIDCONTROLE = MyItens.UNIDCONTROLE
                  CItensNew(CStr(nIDPROD)).VLUNIT = MyItens.VLUNIT
               End If
               MyItens.Rs.MoveNext
            Wend
         End If
      End If
      
      If CItensNew.Count > 0 Then
         .Clear
         Call IncluiGrdLinhaVenda
         
         For Each MyItem In CItensNew
            
            Call IncluiGrdLinhaVenda
            i = .RowCount - 1
            If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
            
            
            Set MyProd = CriarObjeto("BANCO_3R.TB_SPRODUTO")
            Set MyProd.xDb = mvarSys.xDb
            Call MyProd.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPROD:=MyItem.IDPROD)
            If pIDFATURA <> 0 Then
               MyProd.UNIDCONTROLE = MyItem.UNIDCONTROLE
               MyProd.VLVENDA = MyItem.VLUNIT
            End If
         
            Call SetTag(sTag, "IDITEM", i)
            Call SetTag(sTag, "IDPROD", MyItem.IDPROD)
            Call SetTag(sTag, "QTDVENDA", MyItem.QTDVENDA)
            Call SetTag(sTag, "UNIDCONTROLE", MyProd.UNIDCONTROLE)
            Call SetTag(sTag, "VLUNIT", ValBr(MyProd.VLVENDA))
            Call SetTag(sTag, "Total", ValBr(xVal(.CellValue(i, "QTDVENDA")) * xVal(.CellValue(i, "VLUNIT"))))
                              
            .RowTag(i) = sTag
            
            .CellValue(i, "IDITEM") = i
            .CellValue(i, "IDPROD") = MyItem.IDPROD
            .CellValue(i, "ESERVICO") = MyProd.ESERVICO
            .CellValue(i, "QTDVENDA") = MyItem.QTDVENDA
            .CellValue(i, "UNIDCONTROLE") = MyProd.UNIDCONTROLE
            .CellValue(i, "VLUNIT") = ValBr(MyProd.VLVENDA)
            .CellValue(i, "Total") = ValBr(xVal(.CellValue(i, "QTDVENDA")) * xVal(.CellValue(i, "VLUNIT")))
                     
         Next
         
         .RowVisible(1) = True
         '.CellSelected(1, "FOTOTIPO") = True
         .Tag = SetTag(.Tag, "ISDIRT", 0)
      End If
      If .RowCount > 1 Then
         Call SetTag(mvarMe.GrdVenda, "ISDIRT", 1)
      End If
      If Trim(sMsg) <> "" Then
         Call ExibirAviso("Tratamentos sem produto de associado:" & vbNewLine & "(" & Mid(sMsg, 3) & ")")
      End If
   End With
End Sub
Private Function GetSqlGrdAtendimento(Optional pIDATEND As Long, Optional pIDFATURA As Long, Optional bNewVenda As Boolean) As String
   Dim Sql  As String
   Sql = ""
   Sql = Sql & " Select 'Atend.' [TIPO]" & vbNewLine
   'Sql = Sql & ", Right(REPLICATE('0',6)+ Cast(A.IDATENDIMENTO as varchar),6) [IDATENDIMENTO]" & vbNewLine
   Sql = Sql & ", [IDATENDIMENTO]=case A.IDATENDIMENTO When 0 then '--' else Right(REPLICATE('0',6)+ Cast(A.IDATENDIMENTO as varchar),6) end" & vbNewLine
   Sql = Sql & ", [DATA]=case A.IDATENDIMENTO When 0 then Convert(varchar, A.TIMESTAMP, 103) else Convert(varchar, A.DTATEND, 103) end" & vbNewLine
   Sql = Sql & ", [NOME]=case A.IDATENDIMENTO When 0 then 'SISTEMA' else C.NOME end" & vbNewLine
   Sql = Sql & ", 0 [VALOR], Count(S.IDSESSAO) [Itens]" & vbNewLine
   Sql = Sql & ", [DT]=case A.IDATENDIMENTO When 0 then A.TIMESTAMP else A.DTATEND end" & vbNewLine
   Sql = Sql & " From OATENDIMENTO A" & vbNewLine
   Sql = Sql & " Join OCLIENTE C On A.IDLOJA=C.IDLOJA And A.IDCLIENTE=C.IDCLIENTE" & vbNewLine
   If pIDATEND = 0 Then
      Sql = Sql & " Left Join OATENDIMENTO_VENDA AV On AV.IDLOJA=A.IDLOJA And AV.IDATENDIMENTO=A.IDATENDIMENTO" & vbNewLine
      Sql = Sql & " Left Join OSESSAO S On S.IDLOJA=AV.IDLOJA And S.IDATENDIMENTO=AV.IDATENDIMENTO AND S.IDVENDA=AV.IDVENDA And S.IDTPSERVICO<>1" & vbNewLine
      Sql = Sql & " Where A.IDLOJA= " & mvarIDLOJA & vbNewLine
      Sql = Sql & " And AV.IDVENDA = " & mvarIDVENDA & vbNewLine
   Else
      Sql = Sql & " Left Join OSESSAO S On S.IDLOJA=A.IDLOJA And S.IDATENDIMENTO=A.IDATENDIMENTO And S.IDTPSERVICO<>1" & vbNewLine
      If bNewVenda Then
         Sql = Sql & " And (S.IDVENDA=0 Or S.IDVENDA is Null Or S.IDITEM=0 Or S.IDITEM is Null)"
      Else
         Sql = Sql & "      AND S.IDVENDA =" & mvarIDVENDA & vbNewLine
      End If
      Sql = Sql & " Where A.IDLOJA= " & mvarIDLOJA & vbNewLine
      Sql = Sql & " And A.IDATENDIMENTO = " & pIDATEND & vbNewLine
   End If
   Sql = Sql & " Group By A.IDATENDIMENTO, A.DTATEND, C.NOME, A.TIMESTAMP" & vbNewLine
   
   Sql = Sql & " Union All" & vbNewLine
   Sql = Sql & " Select 'Fat.' [TIPO], Right(REPLICATE('0',6)+ Cast(F.IDFATURA as varchar),6) [IDATENDIMENTO]" & vbNewLine
   Sql = Sql & ", Convert(varchar, F.DTEMISSAO, 103) [DATA], C2.NOME, F.VALOR, 0 [Itens]" & vbNewLine
   Sql = Sql & ", F.DTEMISSAO [DT]"
   Sql = Sql & " From FFATURA F" & vbNewLine
   Sql = Sql & " Join OCLIENTE C2 On F.IDLOJA=C2.IDLOJA And F.IDCLIENTE=C2.IDCLIENTE" & vbNewLine
   Sql = Sql & " Where F.IDLOJA=" & mvarIDLOJA & vbNewLine
   If mvarIDVENDA = 0 And mvarIDFATURA <> 0 Then
      Sql = Sql & " And (F.IDVENDA =0 OR F.IDVENDA IS NULL)" & vbNewLine
      Sql = Sql & " And F.IDFATURA= " & mvarIDFATURA & vbNewLine
   Else
      Sql = Sql & " And F.IDVENDA >0" & vbNewLine
      Sql = Sql & " And F.IDVENDA = " & mvarIDVENDA & vbNewLine
   End If
   Sql = Sql & " Order By [DT]"
   
   GetSqlGrdAtendimento = Sql
End Function
Private Sub PopulaGrdAtendimento(Optional pIDATEND As Long, Optional pIDFATURA As Long, Optional bNewVenda As Boolean, Optional bRefreshRow As Boolean)
   Dim MyRs As Object
   Dim sTag As String
   Dim i    As Integer
   Dim Sql  As String

'Call mvarMe_CmdSaldoClick
   
   DoEvents
   Sql = GetSqlGrdAtendimento(pIDATEND, pIDFATURA, bNewVenda)

   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      With mvarMe.GrdAtendimento
         If Not bRefreshRow Then
            .Clear
            Call IncluiGrdLinhaAtendimento(False)
         End If
         While Not MyRs.EOF
            If bRefreshRow Then
               On Error Resume Next
               i = .RowIndex(CStr(MyRs("IDATENDIMENTO") & ""))
               If i <= 0 Then Exit Sub
            Else
               Call IncluiGrdLinhaAtendimento(False)
               i = .RowCount - 1
            End If
            
            If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
            Call SetTag(sTag, "TIPO", MyRs("TIPO") & "")
            Call SetTag(sTag, "IDATENDIMENTO", StrZero(MyRs("IDATENDIMENTO") & "", mvarSys.Propriedades("LENIDVENDA")))
            Call SetTag(sTag, "DATA", MyRs("DATA") & "")
            Call SetTag(sTag, "IDCLIENTE", MyRs("NOME"))
            Call SetTag(sTag, "VALOR", MyRs("VALOR"))
            Call SetTag(sTag, "ITENS", MyRs("ITENS"))
                        
            .RowKey(i) = CStr(MyRs("IDATENDIMENTO") & "")
            .RowTag(i) = sTag
            
            .CellValue(i, "TIPO") = MyRs("TIPO") & ""
            If xVal(MyRs("IDATENDIMENTO") & "") = 0 Then
               .CellValue(i, "IDATENDIMENTO") = MyRs("IDATENDIMENTO") & ""
            Else
               .CellValue(i, "IDATENDIMENTO") = StrZero(MyRs("IDATENDIMENTO") & "", mvarSys.Propriedades("LENIDVENDA"))
            End If
            .CellValue(i, "DATA") = MyRs("DATA") & ""
            .CellValue(i, "IDCLIENTE") = MyRs("NOME") & ""
            .CellValue(i, "VALOR") = MyRs("VALOR") & ""
            arrCellText(i) = MyRs("ITENS") & ""
                     
            MyRs.MoveNext
         Wend
         .RowVisible(1) = True
         .Tag = SetTag(.Tag, "ISDIRT", 0)
      End With
      If pIDATEND <> 0 Then
         Call PopulaNewGRDPRODUTO(pIDFATURA)
         Call SetTag(mvarMe.GrdAtendimento, "ISDIRT", 1)
      End If
   End If
End Sub
Private Sub LimparTela()
   With mvarMe
      .TxtIDVENDA.Text = ""
      .CmdFatura.Caption = ""
      .CmdFatura.Visible = False
      .LblFatura.Visible = False
      
      If mvarIDATENDIMENTO = 0 Then
         .TxtIDVENDA.Enabled = False
         .CmbDTVENDA.Value = mvarSys.xDb.SysDate()
      Else
         Dim TbAtend As Object 'TB_OATENDIMENTO
         
         Set TbAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
         With TbAtend
            Set .xDb = mvarSys.xDb
            mvarMe.TxtIDVENDA.Enabled = False
            If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO) Then
              
               mvarMe.CmbDTVENDA.Value = .DTATEND
               If Format(mvarMe.CmbDTVENDA.Value, "DD/MM/YYYY") <> dToday Then
                  mvarMe.TxtIDVENDA.Enabled = True
               End If
            Else
               mvarMe.CmbDTVENDA.Value = mvarSys.xDb.SysDate()
            End If
            Set TbAtend = Nothing
         End With
      End If
      
      .TxtIDCLIENTE.Text = ""
      Call SetTag(.TxtNOME, "IDCLIENTE", "")
      .TxtTEL1.Text = ""
      .TxtNOME.Text = ""
      
      .TxtVLDESC.Text = ValBr(0)
      .TxtVLPGTO.Text = ValBr(0)
      .TxtVLTROCO.Text = ValBr(0)
      .TxtVLVENDA.Text = ValBr(0)
      .TxtVLSUBTOTAL.Text = ValBr(0)
      .TxtOBS.Text = ""
      If .CmbIDFUNCIONARIO.ListCount > 1 Then .CmbIDFUNCIONARIO.ListIndex = -1
            
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
Private Sub MontaCmbFuncionario()
   Dim TbCmb  As Object 'TB_RFUNCIONARIO
   Dim oCombo As XtremeSuiteControls.ComboBox
   Dim sWhere As String
      
   Set oCombo = mvarMe.CmbIDFUNCIONARIO
   
   Set TbCmb = CriarObjeto("BANCO_3R.TB_RFUNCIONARIO")
   With TbCmb
      Set .xDb = mvarSys.xDb
      oCombo.Clear
      sWhere = ""
      sWhere = sWhere & " SITFUNC='A'"
      sWhere = sWhere & " And FLGVENDA=1"
      
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=sWhere) Then
         While Not .Rs.EOF
            .Popula
            oCombo.AddItem .NOME
            oCombo.ItemData(oCombo.NewIndex) = .IDFUNCIONARIO
            .Rs.MoveNext
         Wend
      Else
         Call ExibirInformacao("Favor cadastrar ao menos um Funcionário.", "Cadastro")
         Dim TlCad As Object
         Set TlCad = CriarObjeto("CADASTRO3R.TL_CADRFUNCIONARIO")
         Set TlCad.Sys = mvarSys
         TlCad.Show
         If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=sWhere) Then
            Call MontaCmbFuncionario
         Else
            Exit Sub
         End If
         Exit Sub
      End If
   End With
   Set TbCmb = Nothing
   With oCombo
      .ListIndex = IIf(.ListCount = 1, 0, -1)
      .Enabled = (.ListCount > 1)
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
         For i = .Count To 1 Step -1
            .Remove i
         Next
         With .Add("IDPROD")
            Sql = "EVENDA=1"
            Sql = Sql & " And SITPROD='A'"
            If TbProd.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=Sql, Ch_ORDERBY:="NMPROD") Then

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
      
      .Clear True
      With .AddCol(sKey:="IDITEM", sHeader:="#", lWidth:=25, bVisible:=True)
         .eType = igCellText
         .sCtrlKey = "IDITEM"
      End With
      With .AddCol(sKey:="IDPROD", sHeader:="Produto", lWidth:=150, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDPROD"
      End With
      With .AddCol(sKey:="ESERVICO", sHeader:="Serviço", lWidth:=65, bVisible:=False)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight
         .sCtrlKey = "ESERVICO"
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
      With .AddCol(sKey:="VLDESC", sHeader:="Desconto", lWidth:=65, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight
         .sCtrlKey = "VLDESC"
      End With
      With .AddCol(sKey:="Total", sHeader:="Total", lWidth:=65, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight
         .sCtrlKey = "Total"
      End With
      With .AddCol(sKey:="DSCDESC", sHeader:="Motivo Desc.", lWidth:=110, bVisible:=False)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight
         .sCtrlKey = "DSCDESC"
      End With
      With .AddCol(sKey:="IDPACOTE", sHeader:="IDPACOTE", lWidth:=25, bVisible:=False)
         .eType = igCellText
         .sCtrlKey = "IDPACOTE"
      End With
            
      For i = 1 To .ColCount
         .ColTag(i) = ""
      Next
            
      .ColTag("IDITEM") = SetTag(.ColTag("IDITEM"), "EDIT", False)
      .ColTag("UNIDCONTROLE") = SetTag(.ColTag("UNIDCONTROLE"), "EDIT", False)
      .ColTag("VLUNIT") = SetTag(.ColTag("VLUNIT"), "EDIT", False)
      .ColTag("VLDESC") = SetTag(.ColTag("VLDESC"), "EDIT", True)
      .ColTag("Total") = SetTag(.ColTag("Total"), "EDIT", False)
      .ColTag("DSCDESC") = SetTag(.ColTag("DSCDESC"), "EDIT", False)
            
      .ColTag("QTDVENDA") = SetTag(.ColTag("QTDVENDA"), "EDITOPT", igTextEditNumberOnly)
      
      .ColTag("IDITEM") = SetTag(.ColTag("IDITEM"), "DEFAULT", "IDENTITY")
      .ColTag("IDPROD") = SetTag(.ColTag("IDPROD"), "DEFAULT", 1)
      .ColTag("UNIDCONTROLE") = SetTag(.ColTag("UNIDCONTROLE"), "DEFAULT", "")
      .ColTag("QTDVENDA") = SetTag(.ColTag("QTDVENDA"), "DEFAULT", 0)
      .ColTag("VLUNIT") = SetTag(.ColTag("VLUNIT"), "DEFAULT", ValBr(0))
      .ColTag("Total") = SetTag(.ColTag("Total"), "DEFAULT", ValBr(0))
      .ColTag("VLDESC") = SetTag(.ColTag("VLDESC"), "DEFAULT", ValBr(0))
      .ColTag("DSCDESC") = SetTag(.ColTag("DSCDESC"), "DEFAULT", "")
      
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
         For i = .Count To 1 Step -1
            .Remove i
         Next
         With .Add("IDFORMAPGTO")
            If TbForma.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="ATIVO=1", Ch_ORDERBY:="IDFORMAPGTO") Then
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
            If TbCond.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDCONDPGTO=1", Ch_ORDERBY:="IDCONDPGTO") Then
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
      
      .Clear True
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
   Dim i       As Integer
   Dim Sql     As String
   Dim sAux    As String
                  
   With mvarMe.GrdAtendimento
      .ShowControlsInAllCells = True
      .Header.DragCols = False
      .SilentValidation = True

      .FocusRect = False
      .GridLines = igGridLinesHorizontal
      .DefaultRowHeight = 17
      
      '.Header.ImageList = mvarMe.IlsIcons
      '.ImageList = mvarMe.IlsIcons
      .Clear True
      With .AddCol(sKey:="TIPO", sHeader:="Doc.", lWidth:=50, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextCenter
         .sCtrlKey = "TIPO"
      End With
           
      With .AddCol(sKey:="IDATENDIMENTO", sHeader:="Nº", lWidth:=50, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft
         .sCtrlKey = "IDATENDIMENTO"
         .oForeColor = vbBlack
         
      End With
      With .AddCol(sKey:="DATA", sHeader:="Emissão", lWidth:=70, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft
         .sCtrlKey = "DATA"
      End With
      With .AddCol(sKey:="IDCLIENTE", sHeader:="Cliente", lWidth:=120, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft
         .sCtrlKey = "IDCLIENTE"
      End With
      With .AddCol(sKey:="VALOR", sHeader:="VALOR", lWidth:=100, bVisible:=False)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft
         .sCtrlKey = "VALOR"
      End With
      With .AddCol(sKey:="ITENS", sHeader:="Itens", lWidth:=60, bVisible:=True)
         '.eType = igCellText
         .eType = igCellCustomDraw
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft
         .sCtrlKey = "ITENS"
      End With
      For i = 1 To .ColCount
         .ColTag(i) = ""
      Next
            
      .ColTag("TIPO") = SetTag(.ColTag("TIPO"), "EDIT", False)
      .ColTag("IDATENDIMENTO") = SetTag(.ColTag("IDATENDIMENTO"), "EDIT", False)
      .ColTag("DATA") = SetTag(.ColTag("DATA"), "EDIT", False)
      .ColTag("IDCLIENTE") = SetTag(.ColTag("IDCLIENTE"), "EDIT", False)
            
      .ColTag("IDATENDIMENTO") = SetTag(.ColTag("IDATENDIMENTO"), "EDITOPT", igTextEditNumberOnly)
      .ColTag("DATA") = SetTag(.ColTag("DATA"), "EDITOPT", igTextEditNumberOnly)
            
            
      IncluiGrdLinhaAtendimento

      Set m_oFont = New StdFont
      m_oFont.Name = .Font.Name
      m_oFont.Size = .Font.Size
      m_oFont.Underline = True
   End With
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
         Call mvarMe_GrdVendaAfterCommitEdit(lRow, .ColIndex("IDPROD"))
         
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
   Dim nVLPgto    As Currency
     
   If Not GetTag(mvarMe, "1VEZ", "1") = "1" Then
      If xVal(mvarMe.TxtVLVENDA.Text) <= xVal(mvarMe.TxtVLPGTO) Then
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
         
         nVLPgto = 0
         For i = 1 To .RowCount
            nVLPgto = nVLPgto + xVal(.CellValue(i, "VLPGTO"))
         Next
         nVlAPagar = xVal(mvarMe.TxtVLVENDA.Text)
         nVlAPagar = nVlAPagar - nVLPgto
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
         Call mvarMe_GrdPagamentoAfterCommitEdit(lRow, .ColIndex("IDFORMAPGTO"))
      
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
   Dim SqlIn   As String
   Dim MYADD   As Object 'DsAdd
   Dim n       As Object
   Dim i       As Integer
   Dim MySel   As Collection
   Dim bRedraw As Boolean
   Dim lRow    As Long
   Dim nColDefault As Integer
   Dim bCliente As Boolean
   Dim sMsg    As String
   
   nColDefault = 2
   With mvarMe.GrdAtendimento
       bRedraw = .Redraw
      '.Header.Font.Size = 9
      .Header.Font.Name = mvarMe.Font.Name
      .Font.Size = .Header.Font.Size
      .Font.Name = .Header.Font.Name
      
      lRow = .RowCount
      If lRow > 0 Then .RemoveRow lRow
      If lRow > 1 Then
         .CellForeColor(lRow - 1, "IDATENDIMENTO") = vbBlack
         '.CellFont(lRow - 1, "IDATENDIMENTO").Underline = True
      End If
      
      'If Not (mvarIDATENDIMENTO = 0 And lRow = 0) Then
      If lRow > 0 Then
         .AddRow
         lRow = .RowCount
         .RowHeight(.RowCount) = .DefaultRowHeight
   
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
         sMsg = "Apenas o cliente corrente? "
         bCliente = (vbYes = ExibirPergunta(sMsg, "Atendimentos / Faturas"))
      
         Sql = "Select Distinct 'Atend.' [TIPO], Right(REPLICATE('0'," + CStr(mvarSys.Propriedades("LENIDVENDA")) + ")+ Cast(A.IDATENDIMENTO as varchar)," + CStr(mvarSys.Propriedades("LENIDVENDA")) + ") [IDATENDIMENTO]" & vbNewLine
         Sql = Sql & " , Convert(varchar, A.DTATEND, 103) [DATA], C.NOME, '--' [VALOR], Count(S.IDSESSAO) [Itens]" & vbNewLine
         Sql = Sql & " From OATENDIMENTO A" & vbNewLine
         Sql = Sql & " Left Join OSESSAO S               On A.IDLOJA=S.IDLOJA And A.IDATENDIMENTO=S.IDATENDIMENTO And S.IDTPSERVICO IN (Select Distinct IDTPSERVICO From OTRATAMENTO_PROD)" & vbNewLine
         Sql = Sql & " Left Join OATENDIMENTO_PRODUTO P  On P.IDLOJA=A.IDLOJA And P.IDATENDIMENTO=A.IDATENDIMENTO" & vbNewLine
         Sql = Sql & " Left Join OCLIENTE C              On A.IDLOJA=C.IDLOJA And A.IDCLIENTE=C.IDCLIENTE" & vbNewLine
         Sql = Sql & " Left Join OATENDIMENTO_VENDA V    On A.IDLOJA=V.IDLOJA And A.IDATENDIMENTO=V.IDATENDIMENTO" & vbNewLine
         Sql = Sql & " Where A.IDLOJA=" & mvarIDLOJA & vbNewLine
         Sql = Sql & " And V.IDVENDA Is Null" & vbNewLine
         Sql = Sql & " And (Not S.IDTPSERVICO Is Null Or Not P.IDPROD Is Null)" & vbNewLine
         If bCliente Then Sql = Sql & " And A.IDCLIENTE=" & mvarIDCLIENTE & vbNewLine
         If mvarMe.GrdAtendimento.RowCount > 1 Then
            SqlIn = ""
            For i = 1 To mvarMe.GrdAtendimento.RowCount - 1
               If mvarMe.GrdAtendimento.CellValue(i, "TIPO") = "Atend." Then
                  SqlIn = SqlIn & IIf(SqlIn = "", "", ", ")
                  SqlIn = SqlIn & Val(mvarMe.GrdAtendimento.CellValue(i, "IDATENDIMENTO"))
               End If
            Next
            If Trim(SqlIn) <> "" Then
               Sql = Sql & " And A.IDATENDIMENTO Not In ( " & SqlIn & " )" & vbNewLine
            End If
         End If
         Sql = Sql & " Group By A.IDATENDIMENTO, A.DTATEND,  C.NOME" & vbNewLine
         Sql = Sql & " Union All" & vbNewLine
         Sql = Sql & " Select Distinct 'Fat.' [TIPO],  Right(REPLICATE('0',6)+ Cast(F.IDFATURA as varchar),6) [IDATENDIMENTO]" & vbNewLine
         Sql = Sql & " , Convert(varchar, F.DTEMISSAO, 103) [DATA]" & vbNewLine
         Sql = Sql & " , C2.NOME, 'R$ ' + replace(cast(F.VALOR as varchar), '.', ',') [VALOR], 0 [Itens]" & vbNewLine
         Sql = Sql & " From FFATURA F" & vbNewLine
         Sql = Sql & " Left Join OCLIENTE C2 On F.IDLOJA=C2.IDLOJA And F.IDCLIENTE=C2.IDCLIENTE" & vbNewLine
         Sql = Sql & " Where F.IDLOJA=" & mvarIDLOJA & vbNewLine
         Sql = Sql & " And (F.IDVENDA Is Null Or F.IDVENDA = 0)" & vbNewLine
         If bCliente Then Sql = Sql & " And F.IDCLIENTE=" & mvarIDCLIENTE
         If mvarMe.GrdAtendimento.RowCount > 1 Then
            SqlIn = ""
            For i = 1 To mvarMe.GrdAtendimento.RowCount - 1
               If mvarMe.GrdAtendimento.CellValue(i, "TIPO") = "Fat." Then
                  SqlIn = SqlIn & IIf(SqlIn = "", "", ", ")
                  SqlIn = SqlIn & Val(mvarMe.GrdAtendimento.CellValue(i, "IDATENDIMENTO"))
               End If
            Next
            If Trim(SqlIn) <> "" Then
               Sql = Sql & " And F.IDFATURA Not In ( " & SqlIn & " )" & vbNewLine
            End If
         
         End If
         Sql = Sql & " Order By IDATENDIMENTO"
         
         'Set MyAdd = New DsAdd
         Set MYADD = CriarObjeto("XActive.DsAdd")
         With MYADD
            .xDb = mvarSys.xDb
            .Caption = "Atendimentos e Faturas Sem Venda."
            .Chaves.Add "TIPO", "TIPO"
            .Chaves.Add "IDATENDIMENTO", "IDATENDIMENTO"
            .Chaves.Add "DATA", "DATA"
            .Chaves.Add "NOME", "NOME"
            .Chaves.Add "VALOR", "VALOR"
            .Chaves.Add "ITENS", "ITENS"
            With .TITULO
               '& 0=eAlin.Esquerda, 1=eAlin.Direita, 2=eAlin.Centro
               Call .Add("TIPO", "Doc.", 8, 0)
               Call .Add("IDATENDIMENTO", "Nº", 7, 0)
               Call .Add("DATA", "Data", 10, 2)
               Call .Add("NOME", "Nome", 25, 0)
               Call .Add("VALOR", "Valor", 10, 1)
               Call .Add("ITENS", "Itens", 5, 2)
            End With
            .CheckBox = True
            .QUERY = Sql
      
            .Show
            Set MySel = .Selecao
         End With
         
         .RemoveRow lRow
         If Not MySel Is Nothing Then
            For Each n In MySel
               .AddRow
               .CellValue(.RowCount, "TIPO") = n("TIPO")
               .CellValue(.RowCount, "IDATENDIMENTO") = n("IDATENDIMENTO")
               .CellValue(.RowCount, "DATA") = n("DATA")
               .CellValue(.RowCount, "IDCLIENTE") = n("NOME")
               .CellValue(.RowCount, "VALOR") = xVal(n("VALOR"))
               arrCellText(.RowCount) = xVal(n("ITENS"))
               
               For i = 1 To .ColCount
                  If GetTag(.ColTag(i), "EDIT", True) = False Then
                     '.CellForeColor(.RowCount, i) = vbGrayText
                  End If
               Next
               
            Next
            If MySel.Count > 0 Then
               Call PopulaGRDPRODUTO
               Call CalculaVenda
               Call PopulaValores
            End If
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
      .Redraw = bRedraw
   End With
End Sub
Private Sub PopulaCliente(Optional pIDCLIENTE As Long, Optional pTelefone As String = "")
   Dim TbCli   As Object 'TB_OCLIENTE
   Dim bAchou  As Boolean
   
   If pIDCLIENTE > 0 Then
      Set TbCli = CriarObjeto("BANCO_3R.TB_OCLIENTE", False)
      With TbCli
         Set .xDb = mvarSys.xDb
         If IsMissing(pIDCLIENTE) Then
            bAchou = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="TEL1='" & pTelefone & "'")
            If Not bAchou Then bAchou = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="TEL2='" & pTelefone & "'")
         Else
            bAchou = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCLIENTE:=pIDCLIENTE)
         End If
         If bAchou Then
            mvarIDCLIENTE = .IDCLIENTE
            Call SetTag(mvarMe.TxtNOME, "IDCLIENTE", .IDCLIENTE)
            mvarMe.TxtIDCLIENTE.Text = .IDCLIENTE
            mvarMe.TxtNOME.Text = .NOME
            mvarMe.TxtTEL1.Text = IIf(.TEL1 = "", .TEL2, .TEL1)
         Else
            mvarIDCLIENTE = 0
            Call SetTag(mvarMe.TxtNOME, "IDCLIENTE", "")
            mvarMe.TxtIDCLIENTE.Text = ""
            mvarMe.TxtNOME.Text = ""
            mvarMe.TxtTEL1.Text = ""
         End If
      End With
   End If
   Set TbCli = Nothing
End Sub
Private Sub mvarMe_TxtIDCLIENTELostFocus()
   If xVal(mvarMe.TxtIDCLIENTE.Text) = 0 Then
      mvarMe.TxtIDCLIENTE.Text = GetTag(mvarMe.TxtNOME, "IDCLIENTE")
   ElseIf xVal(mvarMe.TxtIDCLIENTE.Text) <> xVal(GetTag(mvarMe.TxtNOME, "IDCLIENTE")) Then
      Call PopulaCliente(xVal(mvarMe.TxtIDCLIENTE.Text))
   End If
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
   Sql = Sql & "Select IDTPSERVICO, IDTPTRATAMENTO, IDAREA, IDPROD" & vbNewLine
   Sql = Sql & " From OTRATAMENTO_PROD" & vbNewLine
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA & vbNewLine
   Sql = Sql & " And( isNull(IDTPSERVICO,0)=0 Or IDTPSERVICO =" & pIDTPSERVICO & " )" & vbNewLine
   Sql = Sql & " And ( isNull(IDTPTRATAMENTO,0)=0 Or IDTPTRATAMENTO =" & pIDTPTRATAMENTO & " )" & vbNewLine
   Sql = Sql & " And ( isNull(IDAREA,0)=0 Or IDAREA =" & pIDAREA & " )" & vbNewLine
   Sql = Sql & " Order By IDTPSERVICO DESC, IDTPTRATAMENTO DESC, IDAREA DESC, IDPROD" & vbNewLine
   
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      GetProdutoVenda = Val(MyRs("IDPROD") & "")
   End If
End Function
Private Function CalculaVenda() As Double
   Dim i          As Integer
   Dim nVlDesc     As Currency
   Dim nVLVENDA    As Currency
   Dim nVLSUB      As Currency
   Dim nVLACRESC   As Currency
   
   With mvarMe.GrdVenda
      nVLVENDA = 0
      nVLSUB = 0
      nVlDesc = 0
      For i = 1 To .RowCount - 1
         nVLSUB = nVLSUB + xVal(.CellValue(i, "QTDVENDA")) * xVal(.CellValue(i, "VLUNIT"))
         nVlDesc = nVlDesc + xVal(.CellValue(i, "VLDESC"))
      Next
   End With
   With mvarMe.GrdAtendimento
      For i = 1 To .RowCount
         nVLSUB = nVLSUB + xVal(.CellValue(i, "Valor"))
      Next
   End With
   
   nVLACRESC = xVal(mvarMe.TxtVLACRESC.Text)
   mvarMe.TxtVLSUBTOTAL.Text = ValBr(nVLSUB)
   If mvarMe.Visible = True Then mvarMe.TxtVLDESC.Text = ValBr(nVlDesc)
   
   nVLVENDA = nVLSUB + nVLACRESC - nVlDesc
   mvarMe.TxtVLVENDA.Text = ValBr(nVLVENDA)
   
   CalculaVenda = nVLVENDA
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

Private Sub mvarMe_TxtVLACRESCChange()
   If mvarMe.TxtVLACRESC.Visible Then
      CalculaVenda
      PopulaValores
   'ElseIf mvarIDFATURA = 0 Then
   '   CalculaVenda
   '   Call PopulaValores
   End If
End Sub

Private Sub mvarMe_TxtVLDESCChange()
   If mvarMe.TxtVLDESC.Visible Then
      CalculaVenda
      PopulaValores
'   ElseIf mvarIDFATURA = 0 Then
'      CalculaVenda
'      Call PopulaValores
   End If
End Sub
Private Sub mvarMe_TxtVLPGTOChange()
   If mvarMe.TxtVLDESC.Visible Then
      PopulaValores
   ElseIf mvarIDFATURA = 0 Then
      Call PopulaValores
   End If
End Sub
Private Sub SetExistePromo()
   Dim Sql As String
   Sql = " Select *" & vbNewLine
   Sql = Sql & " From CPROMOCAO" & vbNewLine
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA & vbNewLine
   Sql = Sql & " And ATIVO=1" & vbNewLine
   Sql = Sql & " And DTINIV<=" & SqlDate(mvarSys.xDb.SysDate()) & vbNewLine
   Sql = Sql & " And DTFIMV>=" & SqlDate(mvarSys.xDb.SysDate()) & vbNewLine
   bExistePromo = mvarSys.xDb.ExisteReg(Sql)
End Sub
