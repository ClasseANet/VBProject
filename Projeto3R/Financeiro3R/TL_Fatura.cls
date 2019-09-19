VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Fatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe    As FrmFatura
Attribute mvarMe.VB_VarHelpID = -1
Dim nLenID   As Integer
Dim dToday  As Date
Dim bFaturaPaga As Boolean

Private mvarSys            As Object
Private mvarIDLOJA         As Integer
Private mvarIDATENDIMENTO  As Long
Private mvarIDCLIENTE      As Long
Private mvarIDFATURA       As Long
Private mvarIDVENDA        As Long
Private mvarIDVENDAORIGEM  As Long
Private mvarTlFaturas      As TL_Faturas

Private mvarVLFATURA       As Currency
Public Property Let VLFATURA(ByVal vData As Currency)
    mvarVLFATURA = vData
End Property
Public Property Get VLFATURA() As Currency
   VLFATURA = mvarVLFATURA
End Property
Public Property Set TlFaturas(ByRef vData As Object)
   Set mvarTlFaturas = vData
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
Public Property Let IDVENDAORIGEM(ByVal vData As Long)
    mvarIDVENDAORIGEM = vData
End Property
Public Property Get IDVENDAORIGEM() As Long
   IDVENDAORIGEM = mvarIDVENDAORIGEM
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
   nLenID = mvarSys.Propriedades("LENIDVENDA")
   On Error Resume Next
   dToday = CDate(Format(mvarSys.xDb.SysDate, "DD/MM/YYYY"))
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmFatura
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
End Sub
Public Sub Show(Optional pMode As FormShowConstants = vbModal)
   mvarMe.Show pMode
End Sub
Private Sub mvarMe_Activate()
'   Call MontaTela
'   Call PopulaTela
   If GetTag(mvarMe, "1VEZ", "1") = "1" Then
      Call SetTag(mvarMe, "1VEZ", "0")
      If mvarIDFATURA = 0 And xVal(mvarMe.TxtVALOR) <> 0 Then
         mvarMe.TxtDTPREV.SetFocus
         Call SelecionarTexto(mvarMe.TxtDTPREV)
         mvarMe.Top = mvarMe.Top + 270
         mvarMe.Left = mvarMe.Left + 270
      Else
         mvarMe.TxtVALOR.SetFocus
         Call SelecionarTexto(mvarMe.TxtVALOR)
      End If
      Screen.MousePointer = vbDefault
   End If
End Sub
Private Sub mvarMe_CmbDTPREVChange()
   mvarMe.TxtDTPREV.Text = Format(mvarMe.CmbDTPREV.Value, "dd/mm/yyyy")
End Sub

Private Sub mvarMe_CmdAtendClick()
   Call ExibirAtendimento(mvarMe.TxtIDATEND.Text)
End Sub

Private Sub mvarMe_CmdCancelClick()
   Unload mvarMe
End Sub

Private Sub mvarMe_CmdDividirClick()
   Dim pIDFATURA As Long
   Dim pIDVENDAORIGEM As Long
   Dim pIDATENDIMENTO As Long
   Dim pIDCLIENTE As Long
   Dim pVLFATURA As Currency
'Private Sub ExibirFatura(ByRef pIDFATURA As Long, Optional pIDVENDAORIGEM, Optional pIDATENDIMENTO, Optional pIDCLIENTE As Long, Optional pVLFATURA)
   pIDFATURA = 0
   pIDVENDAORIGEM = mvarIDVENDAORIGEM
   pIDATENDIMENTO = mvarIDATENDIMENTO
   pIDCLIENTE = mvarIDCLIENTE
   pVLFATURA = xVal(InputBox("Entre com novo Valor", "Divisão de Fatura", xVal(mvarMe.TxtVALOR.Text)))
   If pVLFATURA >= xVal(mvarMe.TxtVALOR.Text) Then
      Call ExibirInformacao("Valor inválido", "Fatura")
      GoTo Saida
   End If
   If pVLFATURA > 0 Then
      Dim oNg As NG_Financeiro
      
      On Error GoTo Saida
      Screen.MousePointer = vbHourglass
      
      Set oNg = New NG_Financeiro
      Set oNg.Sys = mvarSys
      
      Call oNg.ExibirFatura(pIDFATURA, pIDVENDAORIGEM, pIDATENDIMENTO, pIDCLIENTE, pVLFATURA)
      If pIDFATURA <> 0 Then
         mvarMe.TxtVALOR.Text = ValBr(xVal(mvarMe.TxtVALOR.Text) - pVLFATURA)
         Call SalvarFatura
      End If
   End If
Saida:
   Screen.MousePointer = vbDefault
'End Sub
End Sub

Private Sub mvarMe_CmdExcluirClick()
   Dim sMsg As String
   Dim Queries As Collection
   Dim Sql As String
   Dim nRecibo As Long
   Dim MyRs As Object
      
   Dim TBFatura As Object 'TB_FFATURA
   
   sMsg = "Esta operação irá excluir definitivamente a fatura '" & mvarMe.TxtIDFATURA.Text & "'"
   sMsg = sMsg & vbNewLine & vbNewLine
   sMsg = sMsg & "Deseja continuar?"
   If ExibirPergunta(sMsg, , False) = vbYes Then
      If SenhaMestre(mvarSys) Then
         Set TBFatura = CriarObjeto("BANCO_3R.TB_FFATURA")
         With TBFatura
            Set .xDb = mvarSys.xDb
            If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDFATURA:=mvarIDFATURA) Then
               If .IDVENDA <> 0 Then
                  Call ExibirInformacao("Fatura já possui venda associada e não pode ser excluída.", "Fatura")
                  Exit Sub
               End If
            End If
         End With
         Set Queries = New Collection
         Queries.Add TBFatura.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDFATURA:=mvarIDFATURA)
         
         Sql = "Update CVENDA "
         Sql = Sql & " Set IDFATURA=0"
         Sql = Sql & ", ALTERSTAMP=1"
         Sql = Sql & ", TIMESTAMP=GetDate()"
         Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
         Sql = Sql & " And IDFATURA=" & mvarIDFATURA
         Queries.Add Sql
         
         If mvarSys.xDb.Executa(Queries) Then
            mvarIDFATURA = 0
            Call ExibirResultado(mvarSys, True)
            If Not mvarTlFaturas Is Nothing Then
               Call mvarTlFaturas.F_Refresh
            End If
            Unload mvarMe
         Else
            Call ExibirResultado(mvarSys, False)
         End If

      Else
         Call ExibirAviso("Esta senha não permite a operação.")
      End If
   End If
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
   DoEvents
   If mvarMe.CmdSalvar.Enabled Then
      If Not VerificaFatura() Then Exit Sub
      If SalvarFatura Then
         Unload mvarMe
      End If
   Else
      Unload mvarMe
   End If
End Sub
Private Function VerificaFatura() As Boolean
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
      Exit Function
   End If
   
   If mvarMe.TxtDTPREV.Text = "" Then
      Call ExibirInformacao("Data da vencimento inválida." & vbNewLine & "Favor verificar data.")
      If mvarMe.TxtDTPREV.Enabled Then
         mvarMe.TxtDTPREV.SetFocus
      End If
      Exit Function
   End If
   If xVal(mvarMe.TxtIDFATURA) = 0 Then
      If CDate(Format(mvarMe.TxtDTPREV.Text, "DD/MM/YYYY")) < dToday Then
         Call ExibirInformacao("Data da vencimento não pode ser menor que hoje." & vbNewLine & "Favor verificar data.")
         If mvarMe.TxtDTPREV.Enabled Then
            mvarMe.TxtDTPREV.SetFocus
         End If
         Exit Function
      ElseIf CDate(Format(mvarMe.TxtDTPREV.Text, "DD/MM/YYYY")) = dToday Then
         If vbNo = ExibirPergunta("A fatura está com seu vencimento para HOJE. Deseja Continuar?", "Data de Vencimento", False) Then
            If mvarMe.TxtDTPREV.Enabled Then
               mvarMe.TxtDTPREV.SetFocus
            End If
            Exit Function
         End If
      End If
      
   End If
   If xVal(mvarMe.TxtVALOR) <= 0 Then
      Call ExibirInformacao("Valor inválido." & vbNewLine & "Favor verificar.")
      mvarMe.TxtVALOR.SetFocus
      Exit Function
   End If
   If mvarVLFATURA > 0 Then
      If mvarVLFATURA < xVal(mvarMe.TxtVALOR) Then
         Call ExibirInformacao("Valor inválido." & vbNewLine & "Esta Fatura não pode ser maior que R$" & ValBr(mvarVLFATURA))
         mvarMe.TxtVALOR.SetFocus
         Exit Function
      End If
   End If
   VerificaFatura = True
End Function
Private Function SalvarFatura() As Boolean
   Dim bResult    As Boolean
   Dim TBFatura   As Object 'TB_FFATURA
   Dim TbItens    As Object
   Dim cQueries   As Collection
   Dim i          As Integer
   '* Cliente
   Call SalvarCliente
   
   '* Fatura
   Call PopulaClasse_ContaPR(TBFatura)
   bResult = True
   If TBFatura.isDirt Then
      bResult = TBFatura.Salvar
      mvarIDFATURA = TBFatura.IDFATURA
   End If
   If bResult Then
      mvarIDFATURA = TBFatura.IDFATURA
      mvarMe.TxtIDFATURA.Text = StrZero(mvarIDFATURA, mvarSys.Propriedades("LENIDVENDA"))
      If mvarVLFATURA > 0 Then mvarVLFATURA = TBFatura.Valor
      mvarMe.CmdExcluir.Enabled = True
      mvarMe.CmdPagar.Enabled = True
      Call ExibirResultado(mvarSys, bResult)
      If Not mvarTlFaturas Is Nothing Then
         Call mvarTlFaturas.PopulaGrdITEM(True)
      End If
   Else
      Call ExibirInformacao("Erro ao gravar fatura.")
   End If
   SalvarFatura = bResult
Saida:
   Set TBFatura = Nothing

End Function
Private Sub SalvarCliente()
   Dim TBCliente As Object 'TB_OCLIENTE
   Dim bClienteOK As Boolean
   
   If Not (mvarMe.TxtNOME.Enabled) Then Exit Sub
   
   Set TBCliente = CriarObjeto("BANCO_3R.TB_OCLIENTE", False)
   With TBCliente
      Set .xDb = mvarSys.xDb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCLIENTE:=mvarIDCLIENTE) Then
         .IDCOLIGADA = mvarSys.Propriedades("IDCOLIGADA")
         .NOME = Trim(mvarMe.TxtNOME.Text)
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
Private Sub PopulaClasse_ContaPR(ByRef pTBFatura As Object)
   Dim sDate   As String
   
   Set pTBFatura = CriarObjeto("BANCO_3R.TB_FFATURA")
   
   With pTBFatura
      .xDb = mvarSys.xDb
      
      Call .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDFATURA:=mvarIDFATURA)
      .IDFATURA = mvarIDFATURA
      .IDLOJA = mvarIDLOJA
      .IDATENDIMENTO = mvarIDATENDIMENTO
      
      .DTPREV = mvarMe.CmbDTPREV.Value
      .IDCLIENTE = mvarIDCLIENTE
      If mvarIDFATURA = 0 Then
         .DTEMISSAO = dToday
         .SITFATURA = 0
      End If
      .Valor = xVal(mvarMe.TxtVALOR.Text)
      .HISTORICO = mvarMe.TxtHISTORICO.Text
      .IDVENDAORIGEM = mvarIDVENDAORIGEM
   End With
End Sub
Private Sub mvarMe_CmdPagarClick()
   Dim TlVenda As TL_Venda
   Dim Sql As String
   
   Set TlVenda = New TL_Venda
   With TlVenda
      Set .Sys = mvarSys
      .IDLOJA = mvarIDLOJA
      .IDATENDIMENTO = mvarIDATENDIMENTO
      .IDCLIENTE = mvarIDCLIENTE
      .IDFATURA = mvarIDFATURA
      'If mvarIDATENDIMENTO = 0 Then
      '   .IDVENDA = mvarIDVENDAORIGEM
      'Else
         .IDVENDA = mvarIDVENDA
      'End If
      
      .Show vbModal
      
'      If .IDVENDA <> 0 And mvarIDVENDA = 0 Then
         mvarIDVENDA = .IDVENDA
'         Sql = "Update FFATURA "
'         Sql = Sql & " Set SITFATURA=10"
'         Sql = Sql & " , IDVENDA=" & SqlNum(.IDVENDA)
'         Sql = Sql & " Where IDLOJA= " & SqlNum(mvarIDLOJA)
'         Sql = Sql & " And IDFATURA= " & SqlNum(mvarIDFATURA)
'         mvarsys.xdb.Executa Sql
'      End If
   End With
   Set TlVenda = Nothing
   
   If mvarIDVENDA > 0 Then
      Call PopulaTela
   End If
End Sub

Private Sub mvarMe_CmdSalvarClick()
   DoEvents
   If Not VerificaFatura() Then Exit Sub
   Call SalvarFatura
End Sub
Private Sub mvarMe_CmdVendaClick()
   Dim MyNg As NG_Financeiro
   Set MyNg = New NG_Financeiro
   Set MyNg.Sys = mvarSys
   Call MyNg.ExibirVenda(mvarIDVENDAORIGEM)
   Set MyNg = Nothing
End Sub
Private Sub mvarMe_CmdVendaPClick()
   Dim MyNg As NG_Financeiro
   Set MyNg = New NG_Financeiro
   Set MyNg.Sys = mvarSys
   Call MyNg.ExibirVenda(mvarIDVENDA)
   Set MyNg = Nothing
End Sub

Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim lRow As Long
   
   If KeyCode = vbKeyDelete Then
     
   ElseIf KeyCode = vbKeyF9 Then
      If Not bFaturaPaga Then
         Call mvarMe_CmdPagarClick
      End If
   ElseIf KeyCode = vbKeyEscape Then
      Call mvarMe_CmdCancelClick
   End If
End Sub
Private Sub mvarMe_Load()
   mvarMe.Icon = mvarSys.MDI.Icon
   '* Monta Tela
   Call MontaTela
   Call PopulaTela
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub MontaTela()
   With mvarMe
      .CmdExcluir.Enabled = False
      .CmdPagar.Enabled = False
      .TxtDTEMISSAO.Enabled = False
      .CmbDTEMISSAO.Enabled = False
      .TxtIDFATURA.BackColor = &HCCEEDD
      .CmdDividir.Enabled = False
   End With
End Sub
Private Sub PopulaTela()
   Dim Sql     As String
   Dim TBFatura As Object 'TB_FFATURA
   Dim TBAtend  As Object 'TB_FFATURA
   Dim bOk     As Boolean
   Dim sAux    As String
   
   Call LimparTela
   
   Set TBFatura = CriarObjeto("BANCO_3R.TB_FFATURA")
   Set TBFatura.xDb = mvarSys.xDb
'   If mvarIDFATURA = 0 And mvarIDATENDIMENTO <> 0 Then
'      If TBFatura.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDATENDIMENTO=" & mvarIDATENDIMENTO) Then
'         mvarIDFATURA = TBFatura.IDFATURA
'      End If
'   End If

'   If mvarIDFATURA = 0 And mvarIDVENDAORIGEM <> 0 And mvarVLFATURA > 0 Then
'      If TBFatura.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDVENDAORIGEM=" & mvarIDVENDAORIGEM) Then
'         mvarIDFATURA = TBFatura.IDFATURA
'      End If
'   End If
   
   If mvarIDFATURA > 0 Then
      mvarMe.TxtDTPREV.Enabled = True
      mvarMe.CmbDTPREV.Enabled = True
      mvarMe.CmdPagar.Enabled = True
      mvarMe.CmdExcluir.Enabled = True
      mvarMe.CmdDividir.Enabled = True
      
      If TBFatura Is Nothing Then
         Set TBFatura = CriarObjeto("BANCO_3R.TB_FFATURA")
         Set TBFatura.xDb = mvarSys.xDb
      End If
      If TBFatura.IDFATURA = mvarIDFATURA Then
         bOk = True
      Else
         bOk = TBFatura.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDFATURA:=mvarIDFATURA)
      End If
         
      If bOk Then
         With TBFatura
            mvarMe.CmbDTEMISSAO.Enabled = (.IDATENDIMENTO = 0)
            mvarMe.TxtDTEMISSAO.Enabled = (.IDATENDIMENTO = 0)
                     
            mvarMe.TxtIDFATURA.Text = StrZero(.IDFATURA, mvarSys.Propriedades("LENIDVENDA"))
            mvarMe.CmbDTEMISSAO.Value = CDate(.DTEMISSAO)
            mvarMe.TxtDTEMISSAO.Text = .DTEMISSAO
            mvarMe.Caption = " Fatura Nº" & mvarMe.TxtIDFATURA.Text
            mvarMe.CmbDTPREV.Value = CDate(.DTPREV)
            mvarMe.TxtDTPREV.Text = .DTPREV
                                    
            Call PopulaCliente(.IDCLIENTE)
            Call PopulaVenda(.IDVENDA, .IDVENDAORIGEM)

            mvarMe.TxtVALOR.Text = ValBr(.Valor)
            mvarMe.TxtHISTORICO.Text = .HISTORICO
            
            mvarIDATENDIMENTO = .IDATENDIMENTO
            mvarIDVENDA = .IDVENDA
            mvarIDVENDAORIGEM = .IDVENDAORIGEM
            
            If .SITFATURA = 10 Then
               mvarMe.LblSITFAT.Caption = IIf(.SITFATURA = 0, "Em Aberto", "Fechada")
               mvarMe.CmdPagar.Enabled = False
               mvarMe.CmdExcluir.Enabled = False
               mvarMe.CmdSalvar.Enabled = False
               mvarMe.TxtDTPREV.Enabled = False
               mvarMe.CmbDTPREV.Enabled = False
               mvarMe.TxtVALOR.Locked = True
                           
            Else
               mvarMe.LblSITFAT.Caption = IIf(.SITFATURA = 0, "Em Aberto", "Em Aberto")
               mvarMe.CmdPagar.Enabled = True
               mvarMe.CmdExcluir.Enabled = True
               mvarMe.CmdSalvar.Enabled = True
               mvarMe.TxtDTPREV.Enabled = True
               mvarMe.CmbDTPREV.Enabled = True
               mvarMe.TxtVALOR.Locked = False
            End If
            
            mvarMe.CmdAtend.Enabled = False
            Set TBAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
            Set TBAtend.xDb = mvarSys.xDb
            If TBAtend.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO) Then
               mvarMe.TxtIDATEND.Text = StrZero(TBAtend.IDATENDIMENTO, nLenID)
               mvarMe.TxtAtend.Text = Format(TBAtend.DTATEND, "dd/mm/yyyy") & " " & Format(TBAtend.HHINI, "hh:mm")
               mvarMe.CmdAtend.Enabled = True
            End If
         End With
      End If
      Set TBFatura = Nothing
   Else
      Call PopulaCliente(mvarIDCLIENTE)
      
      mvarMe.CmdAtend.Enabled = False
      Set TBAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
      Set TBAtend.xDb = mvarSys.xDb
      If TBAtend.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO) Then
         mvarMe.CmbDTEMISSAO.Enabled = (mvarIDATENDIMENTO = 0)
         mvarMe.TxtDTEMISSAO.Enabled = (mvarIDATENDIMENTO = 0)
         sAux = Format(TBAtend.DTATEND, "dd/mm/yyyy") & " " & Format(TBAtend.HHINI, "hh:mm")
         mvarMe.TxtIDATEND.Text = StrZero(mvarIDATENDIMENTO, nLenID)
         mvarMe.TxtAtend.Text = Format(sAux, "dd/mm/yyyy hh:mm")
         mvarMe.CmdAtend.Enabled = True
         
         Sql = " Select SUM(VLVENDA) [VALOR]"
         Sql = Sql & " From ("
         Sql = Sql & " Select AP.IDATENDIMENTO, 0 [IDSESSAO], IsNull(P1.VLVENDA,0) [VLVENDA]"
         Sql = Sql & " From OATENDIMENTO_PRODUTO AP"
         Sql = Sql & " Join SPRODUTO P1 On AP.IDLOJA=P1.IDLOJA And AP.IDPROD=P1.IDPROD"
         Sql = Sql & " Where AP.IDLOJA=" & SqlNum(mvarIDLOJA)
         Sql = Sql & " And AP.IDATENDIMENTO=" & SqlNum(mvarIDATENDIMENTO)
         Sql = Sql & " Union All"
         
         Sql = Sql & " Select DISTINCT S.IDATENDIMENTO, S.IDSESSAO, IsNull(P2.VLVENDA,0) [VLVENDA]"
         Sql = Sql & " From OSESSAO S"
         Sql = Sql & " Left Join OTRATAMENTO_PROD T On S.IDLOJA=T.IDLOJA AND T.IDPROD=("
         Sql = Sql & " SELECT TOP 1 T2.IDPROD FROM OTRATAMENTO_PROD T2"
         Sql = Sql & " Where (S.IDTPSERVICO = T2.IDTPSERVICO Or T2.IDTPSERVICO Is Null)"
         Sql = Sql & " And (S.IDTPTRATAMENTO=T2.IDTPTRATAMENTO OR T2.IDTPTRATAMENTO Is Null)"
         Sql = Sql & " And (S.IDAREA=T2.IDAREA OR T2.IDAREA Is Null)"
         Sql = Sql & " And S.IDLOJA=T2.IDLOJA"
         Sql = Sql & " And T2.IDLOJA=" & SqlNum(mvarIDLOJA)
         Sql = Sql & " Order By T2.IDTPSERVICO DESC, T2.IDTPTRATAMENTO DESC, T2.IDAREA DESC, T2.IDPROD"
         Sql = Sql & " )"
         Sql = Sql & " Left Join SPRODUTO P2 On T.IDLOJA=P2.IDLOJA And T.IDPROD=P2.IDPROD"
         Sql = Sql & " Where S.IDLOJA=" & SqlNum(mvarIDLOJA)
         Sql = Sql & " And S.IDATENDIMENTO=" & SqlNum(mvarIDATENDIMENTO)
         Sql = Sql & " ) RS"
         If mvarSys.xDb.AbreTabela(Sql) Then
            If xVal(mvarSys.xDb.RsAux("VALOR") & "") = 0 Then
               If mvarVLFATURA <> 0 Then
                  mvarMe.TxtVALOR.Text = ValBr(mvarVLFATURA)
               End If
            Else
               mvarMe.TxtVALOR.Text = ValBr(xVal(mvarSys.xDb.RsAux("VALOR") & ""))
            End If
         End If
         Call PopulaVenda(mvarIDVENDA, mvarIDVENDAORIGEM)
      
      Else
         mvarMe.CmbDTEMISSAO.Enabled = (mvarIDATENDIMENTO = 0)
         mvarMe.TxtDTEMISSAO.Enabled = (mvarIDATENDIMENTO = 0)
         sAux = ""
         mvarMe.TxtIDATEND.Text = ""
         mvarMe.TxtAtend.Text = Format(sAux, "dd/mm/yyyy hh:mm")
         mvarMe.CmdAtend.Enabled = True
         If mvarVLFATURA <> 0 Then
            mvarMe.TxtVALOR.Text = ValBr(mvarVLFATURA)
         End If
      End If
      mvarMe.CmbDTEMISSAO.Value = CDate(dToday)
      mvarMe.TxtDTEMISSAO.Text = dToday
   End If
End Sub
Private Sub LimparTela()
   With mvarMe
      mvarMe.Caption = " Fatura"
      .CmbDTEMISSAO.Value = mvarSys.xDb.SysDate()
      .TxtDTEMISSAO.Text = .CmbDTEMISSAO.Value
      .TxtIDFATURA.Text = ""
      .TxtIDFATURA.Enabled = False
      .CmbDTPREV.Value = mvarSys.xDb.SysDate()
      .TxtDTPREV.Text = .CmbDTPREV.Value
      
      Call SetTag(.TxtNOME, "IDCLIENTE", "")
      .TxtNOME.Text = ""
      
      .TxtVALOR.Text = ValBr(0)
      .TxtHISTORICO.Text = ""
      
      .TxtIDATEND.Text = ""
      .TxtAtend.Text = ""
      .CmdAtend.Enabled = False
      
      .TxtIDVENDAP.Text = ""
      .TxtDTVENDAP.Text = ""
      .CmdVendaP.Enabled = False
      
      .TxtIDVENDA.Text = ""
      .TxtDTVENDA.Text = ""
      .CmdVenda.Enabled = False
      
   End With
End Sub
Private Sub PopulaCliente(pIDCLIENTE As Long)
   Dim TbCli   As Object 'TB_OCLIENTE
   Dim bAchou  As Boolean
   
   Set TbCli = CriarObjeto("BANCO_3R.TB_OCLIENTE", False)
   With TbCli
      Set .xDb = mvarSys.xDb
      bAchou = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCLIENTE:=pIDCLIENTE)
      If bAchou Then
         mvarIDCLIENTE = .IDCLIENTE
         Call SetTag(mvarMe.TxtNOME, "IDCLIENTE", .IDCLIENTE)
         mvarMe.TxtIDCLIENTE.Text = .IDCLIENTE
         mvarMe.TxtNOME.Text = .NOME
      Else
         mvarIDCLIENTE = 0
         Call SetTag(mvarMe.TxtNOME, "IDCLIENTE", "")
         mvarMe.TxtIDCLIENTE.Text = ""
         mvarMe.TxtNOME.Text = ""
      End If
   End With
   Set TbCli = Nothing
End Sub
Private Sub PopulaVenda(pIDVENDA As Long, pIDVENDAORIGEM As Long)
   Dim TBVenda   As Object 'TB_OVENDA
   Dim bAchou  As Boolean
      
   Set TBVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
   With TBVenda
      Set .xDb = mvarSys.xDb
      bAchou = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=pIDVENDA)
      If bAchou Then
         mvarIDVENDA = .IDVENDA
         mvarMe.TxtIDVENDAP.Text = StrZero(.IDVENDA, mvarSys.Propriedades("LENIDVENDA"))
         mvarMe.TxtDTVENDAP.Text = Format(.DTVENDA, "dd/mm/yyyy hh:mm")
         mvarMe.CmdVendaP.Enabled = True
      Else
         mvarIDVENDA = 0
         mvarMe.TxtIDVENDAP.Text = ""
         mvarMe.TxtDTVENDAP.Text = ""
         mvarMe.CmdVendaP.Enabled = False
      End If
   End With
   
   'Set TBVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
   With TBVenda
      Set .xDb = mvarSys.xDb
      bAchou = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=pIDVENDAORIGEM)
      If bAchou Then
         mvarIDVENDAORIGEM = .IDVENDA
         mvarMe.TxtIDVENDA.Text = StrZero(.IDVENDA, mvarSys.Propriedades("LENIDVENDA"))
         mvarMe.TxtDTVENDA.Text = Format(.DTVENDA, "dd/mm/yyyy hh:mm")
         mvarMe.CmdVenda.Enabled = True
      Else
         mvarIDVENDAORIGEM = 0
         mvarMe.TxtIDVENDA.Text = ""
         mvarMe.TxtDTVENDA.Text = ""
         mvarMe.CmdVenda.Enabled = False
      End If
   End With
   Set TBVenda = Nothing
End Sub
Private Sub mvarMe_TxtDTPREVLostFocus()
   mvarMe.TxtDTPREV.Text = FormatarData(mvarMe.TxtDTPREV.Text)
   mvarMe.CmbDTPREV.Value = mvarMe.TxtDTPREV.Text
End Sub
Private Sub mvarMe_TxtIDCLIENTELostFocus()
   If xVal(mvarMe.TxtIDCLIENTE.Text) = 0 Then
      mvarMe.TxtIDCLIENTE.Text = GetTag(mvarMe.TxtNOME, "IDCLIENTE")
   ElseIf xVal(mvarMe.TxtIDCLIENTE.Text) <> xVal(GetTag(mvarMe.TxtNOME, "IDCLIENTE")) Then
      Call PopulaCliente(xVal(mvarMe.TxtIDCLIENTE.Text))
   End If
End Sub
Private Sub mvarMe_TxtIDFATURALostFocus()
   If Val(mvarMe.TxtIDFATURA.Text) = 0 Then
      mvarMe.TxtIDFATURA.Text = ""
   Else
      mvarMe.TxtIDFATURA.Text = StrZero(mvarMe.TxtIDFATURA.Text, mvarSys.Propriedades("LENIDVENDA"))
   End If
End Sub
Private Function GetProdutoVenda(pIDTPSERVICO As Long, pIDTPTRATAMENTO As Long, pIDAREA As Long) As Long
   Dim Sql As String
   Dim MyRs As Object
   
   Sql = ""
   Sql = Sql & "Select IDTPSERVICO, IDTPTRATAMENTO, IDAREA, IDPROD"
   Sql = Sql & " From OTRATAMENTO_PROD"
   Sql = Sql & " Where IDLOJA =" & mvarIDLOJA
   Sql = Sql & " Where ( IDTPSERVICO is Null Or IDTPSERVICO =" & pIDTPSERVICO & " )"
   Sql = Sql & " And ( IDTPTRATAMENTO is Null Or IDTPTRATAMENTO =" & pIDTPTRATAMENTO & " )"
   Sql = Sql & " And ( IDAREA is Null Or IDAREA =" & pIDAREA & " )"
   Sql = Sql & " Order By IDTPSERVICO DESC, IDTPTRATAMENTO DESC, IDAREA DESC, IDPROD"
   
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      GetProdutoVenda = Val(MyRs("IDPROD") & "")
   End If
End Function
Private Sub ExibirAtendimento(pIDATEND As Integer)
   Dim TlAtend As Object
   
   Set TlAtend = CriarObjeto("Calendario3R.TL_Atendimento")
   With TlAtend
      Set .Sys = mvarSys
      .IDLOJA = mvarIDLOJA
      .IDATENDIMENTO = pIDATEND
      .Show vbModal
   End With
   Set TlAtend = Nothing
End Sub

