VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADOMAQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe  As FrmCADOMAQ
Attribute mvarMe.VB_VarHelpID = -1

Private mvarShortBar As TL_ShortBar
'Private WithEvents mvarMDI As FrmMDI

Private mvarRptRec         As ReportRow
Private TbCad             As Object 'TB_OMAQUINA

Private mvarIDLOJA         As Integer
Private mvarIDMAQUINA      As Integer

Private mvarSys            As Object
Public Property Set TLCadastro(ByRef vData As Object)
   Dim sAux As String
   With vData
      .Titulo = "Máquinas"
      
      sAux = "Select IDMAQUINA, IDMAQUINA [Id.], CODMAQUINA [Código] "  ', NREGISTRO [Registro], DISPAROS [Disparos]"
      sAux = sAux & " From OMAQUINA"
      sAux = sAux & " Where IDLOJA = " & mvarIDLOJA
      .Consulta = sAux
      
      sAux = ""
      sAux = sAux & "|CAMPO=IDMAQUINA    |CHAVE=S|GRUPO=0|VISIBLE=N:"
      .TAGCAMPOS = sAux
   End With
End Property
Public Property Set ShortBar(ByVal vData As TL_ShortBar)
   Set mvarShortBar = vData
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = Val(mvarSys.Propriedades("IDLOJA"))
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property

Public Property Let IDLOJA(ByVal vData As String)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As String
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDMAQUINA(ByVal vData As String)
   mvarIDMAQUINA = vData
End Property
Public Property Get IDMAQUINA() As String
   IDMAQUINA = mvarIDMAQUINA
End Property
Public Property Set RptRec(ByVal vData As ReportRow)
   Set mvarRptRec = vData
'   mvarIDLOJA = RptRec.Record.Item(0).Value
   mvarIDMAQUINA = RptRec.Record.Item(1).Value
End Property
Public Property Get RptRec() As ReportRow
   Set RptRec = mvarRptRec
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmCADOMAQ
End Sub
Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call SetTag(mvarMe, "1VEZ", 0)
      'Call MontaTela
      'Call PopulaTela
   End If
End Sub
Private Sub InstanciaClasse()
   Set TbCad = Nothing
   Set TbCad = CriarObjeto("BANCO_3R.TB_OMAQUINA")
   Set TbCad.Xdb = mvarSys.Xdb
End Sub
Private Sub PopulaTela()
   If TbCad Is Nothing Then
      Call InstanciaClasse
   End If
   
   With TbCad
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDMAQUINA:=mvarIDMAQUINA) Then
         mvarMe.CmdExcluir.Enabled = True
         
         mvarMe.TxtID.Text = .IDMAQUINA
         mvarMe.OptATIVO(.SITMAQUINA).Value = True
         mvarMe.TxtCODMAQUINA.Text = .CODMAQUINA
         mvarMe.TxtDTOPERACAO.Text = .DTOPERACAO
         
         Dim oTbTipo As Object
         Set oTbTipo = Nothing
         Set oTbTipo = CriarObjeto("BANCO_3R.TB_OTPMAQ")
         Set oTbTipo.Xdb = mvarSys.Xdb
         If oTbTipo.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDTPMAQ:=.IDTPMAQ) Then
            mvarMe.CmbIDTPMAQ.ListIndex = mvarMe.CmbIDTPMAQ.FindItem(0, oTbTipo.DSCMAQ, True)
         End If
         
         Dim oTbSLMaq As Object
         Set oTbSLMaq = Nothing
         Set oTbSLMaq = CriarObjeto("BANCO_3R.TB_OSALA_MAQUINA")
         Set oTbSLMaq.Xdb = mvarSys.Xdb
         If oTbSLMaq.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDMAQUINA:=mvarIDMAQUINA, Ch_Where:="(DTFIM>=GetDate() Or DTFIM is Null)") Then
            Dim oTbSala As Object
            Set oTbSala = Nothing
            Set oTbSala = CriarObjeto("BANCO_3R.TB_OSALA")
            Set oTbSala.Xdb = mvarSys.Xdb
            If oTbSala.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDSALA:=oTbSLMaq.IDSALA) Then
               mvarMe.CmbIDSALA.ListIndex = mvarMe.CmbIDSALA.FindItem(0, oTbSala.CODSALA, True)
            End If
         End If
         
      Else
         Call LimpaTela
      End If
   End With
End Sub
Public Function Excluir() As Boolean
   Dim Queries As Collection
   
   
   If mvarIDMAQUINA <> 0 Then
      If VerificarDependencias() Then
         If ExibirPergunta("Confirma a exclusão da máquina?", "Atenção", False) = vbYes Then
            Call TbCad.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDMAQUINA:=xVal(mvarMe.TxtID))
            
            Set Queries = New Collection
            Call ExcluirDependencias(Queries)
            Queries.Add TbCad.QryDelete
            
            If mvarSys.Xdb.Executa(Queries) Then
               Excluir = True
               Call ExibirResultado(mvarSys, True)
               
               'mvarIsDirt = True
               
               Call LimpaTela
               If Not mvarShortBar Is Nothing Then
                  mvarShortBar.TLCadastro.F_Refresh
               End If
               If mvarMe.TxtCODMAQUINA.Visible And mvarMe.TxtCODMAQUINA.Enabled Then
                  mvarMe.TxtCODMAQUINA.SetFocus
               End If
            Else
               Call ExibirResultado(mvarSys, False)
            End If
            Set Queries = Nothing
         End If
      End If
   End If
End Function
Private Sub ExcluirDependencias(ByRef pQueries As Collection)
   Dim Tb01 As Object
   
   Set Tb01 = Nothing
   Set Tb01 = CriarObjeto("BANCO_3R.TB_OSALA_MAQUINA")
   Set Tb01.Xdb = mvarSys.Xdb
            
   pQueries.Add Tb01.QryDelete(Ch_IDMAQUINA:=SqlNum(mvarMe.TxtID))
End Sub
Private Function ValidaCampos() As Boolean
   ' *******************
   Dim bResult As Boolean
   
   bResult = True
   If Trim(mvarMe.TxtCODMAQUINA.Text) = "" Then
      ExibirAviso "O código deve ser preenchido."
      bResult = False
      mvarMe.TxtCODMAQUINA.SetFocus
      GoTo Fim
   End If
   If Trim(mvarMe.TxtDTOPERACAO.Text) = "" Then
      ExibirAviso "Data inválida."
      bResult = False
      mvarMe.TxtDTOPERACAO.SetFocus
      GoTo Fim
   End If
   If Trim(mvarMe.CmbIDTPMAQ.Text) = "" Then
      ExibirAviso "Tipo de Máquina Inválido."
      bResult = False
      mvarMe.CmbIDTPMAQ.SetFocus
      GoTo Fim
   End If
   If Trim(mvarMe.CmbIDSALA.Text) = "" Then
      ExibirAviso "Sala Inválida."
      bResult = False
      mvarMe.CmbIDSALA.SetFocus
      GoTo Fim
   End If
   
   
Fim:
   ValidaCampos = bResult
End Function
Private Sub LimpaTela()
   Dim n As Object
   On Error Resume Next
   With mvarMe
      For Each n In mvarMe.Controls
         n.Text = ""
      Next
      .OptATIVO(1).Value = True
      .CmbIDTPMAQ.ListIndex = -1
      .CmbIDSALA.ListIndex = -1
      If .CmbIDTPMAQ.ListCount > 3 Then
         .CmbIDTPMAQ.ListIndex = 3
         .CmbIDTPMAQ.Text = .CmbIDTPMAQ.List(3)
      End If
      If .CmbIDSALA.ListCount > 0 Then
         .CmbIDSALA.ListIndex = 0
         .CmbIDSALA.Text = .CmbIDSALA.List(0)
      End If
      .CmbIDSALA.Enabled = (.CmbIDSALA.Text = "")
   End With
End Sub
Private Sub MontaTela()
   
   Dim MyTb As Object
   Set MyTb = CriarObjeto("BANCO_3R.TB_OTPMAQ")
   Set MyTb.Xdb = mvarSys.Xdb
   If MyTb.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDTPMAQ>=0", Ch_OrderBy:="DSCMAQ") Then
      With mvarMe.CmbIDTPMAQ
         .Clear
         While Not MyTb.Rs.EOF
            MyTb.Popula
            .AddItem MyTb.DSCMAQ, "k" & MyTb.IDTPMAQ
            .ItemData(.NewIndex) = MyTb.IDTPMAQ
            MyTb.Rs.MoveNext
         Wend
         If .ListCount > 0 Then .ListIndex = 0
         If .ListCount > 3 Then .ListIndex = 3
      End With
   End If
   
   Set MyTb = CriarObjeto("BANCO_3R.TB_OSALA")
   Set MyTb.Xdb = mvarSys.Xdb
   If MyTb.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_OrderBy:="CODSALA") Then
      With mvarMe.CmbIDSALA
         .Clear
         While Not MyTb.Rs.EOF
            MyTb.Popula
            .AddItem MyTb.CODSALA, "k" & MyTb.IDSALA
            .ItemData(.NewIndex) = MyTb.IDSALA
            MyTb.Rs.MoveNext
         Wend
         If .ListCount > 0 Then .ListIndex = 0
         If .ListCount = 1 Then .Enabled = False
      End With
   End If
End Sub
Private Sub PopulaClasse(ByRef Queries As Collection)
   Dim oTbSLMaq As Object
   
   If TbCad Is Nothing Then Call InstanciaClasse
   
   With TbCad
      If .IDMAQUINA = 0 Then
'         .DTCADASTRO = mvarSys.Xdb.SysDate()
      End If
      .IDLOJA = mvarIDLOJA
      .CODMAQUINA = mvarMe.TxtCODMAQUINA.Text
      .DTOPERACAO = mvarMe.TxtDTOPERACAO.Text
      .IDTPMAQ = mvarMe.CmbIDTPMAQ.ItemData(mvarMe.CmbIDTPMAQ.ListIndex)
      .SITMAQUINA = IIf(mvarMe.OptATIVO.Item(1).Value = True, 1, 0)
      
      If .isDirt Then
         If .Salvar Then
            mvarIDMAQUINA = .IDMAQUINA
         End If
      End If
      Set oTbSLMaq = Nothing
      Set oTbSLMaq = CriarObjeto("BANCO_3R.TB_OSALA_MAQUINA")
      Set oTbSLMaq.Xdb = mvarSys.Xdb
      Call oTbSLMaq.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDMAQUINA:=mvarIDMAQUINA, Ch_Where:="(DTFIM>=GetDate() Or DTFIM is Null)")
      oTbSLMaq.IDMAQUINA = mvarIDMAQUINA
      oTbSLMaq.IDLOJA = mvarIDLOJA
      oTbSLMaq.IDSALA = mvarMe.CmbIDSALA.ItemData(mvarMe.CmbIDSALA.ListIndex)
      oTbSLMaq.DTINICIO = mvarMe.TxtDTOPERACAO.Text
      If oTbSLMaq.isDirt Then Queries.Add oTbSLMaq.QrySave
      
   End With
End Sub
Private Function Salvar() As Boolean
   Dim bOk     As Boolean
   Dim xNode   As TreeViewNode
   Dim bIsDirt As Boolean
   Dim Queries As Collection
   
   If ValidaCampos() Then
      Set Queries = New Collection
      Call PopulaClasse(Queries)
      
      'bIsDirt = TbCad.isDirt
      'If bIsDirt Then
      '   If TbCad.Salvar() Then
      '      bOk = True
      '   End If
      'Else
      '   bOk = True
      'End If
      bOk = False
      If Queries.Count = 0 Then
         bOk = True
      Else
         If mvarSys.Xdb.Executa(Queries) Then
            bOk = True
         End If
      End If
      
      Call ExibirResultado(mvarSys, bOk, IIf(bOk, 1, 2))
   End If
   Salvar = bOk
End Function
Private Function VerificarDependencias()
   Dim bResult    As Boolean
   Dim TbAtend    As Object
   Dim sMsg       As String
   
   Set TbAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
   With TbAtend
      Set .Xdb = mvarSys.Xdb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDMAQUINA=" & SqlNum(mvarMe.TxtID)) Then
         sMsg = "Este item já possui atendimento associado e não pode ser excluído."
         Call ExibirAviso(sMsg)
         GoTo Saida
      End If
   End With
   
   bResult = True
   
   VerificarDependencias = bResult
Saida:
   Set TbAtend = Nothing
End Function

Private Sub mvarMe_CmbIDTPMAQClick()
   Dim Sql As String
   
   'If xVal(mvarMe.TxtID.Text) = 0 Then
      Sql = "Select Count(*)"
      Sql = Sql & " From OMAQUINA"
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And IDTPMAQ=" & mvarMe.CmbIDTPMAQ.ItemData(mvarMe.CmbIDTPMAQ.ListIndex)
      If xVal(mvarMe.TxtID.Text) <> 0 Then
         Sql = Sql & " And IDMAQUINA<" & SqlNum(xVal(mvarMe.TxtID.Text))
      End If
      If mvarSys.Xdb.AbreTabela(Sql) Then
         mvarMe.TxtCODMAQUINA.Text = mvarMe.CmbIDTPMAQ.Text & "_" & StrZero(xVal(mvarSys.Xdb.RsAux(0)) + 1, 2)
         mvarMe.TxtCODMAQUINA.Enabled = (mvarMe.TxtCODMAQUINA.Text = "")
      End If
   'End If
End Sub

Private Sub mvarMe_CmdExcluirClick()
   If mvarIDMAQUINA <> 0 Then
      If mvarMe.Visible Then
         Call PopulaTela
      End If
      Call Excluir
   End If
End Sub
Private Sub mvarMe_CmdNovoClick()
   If xVal(mvarMe.TxtID) = 0 And Trim(mvarMe.TxtCODMAQUINA) = "" Then
      Call LimpaTela
      mvarMe.TxtCODMAQUINA.SetFocus
   Else
      If Salvar Then
         Call LimpaTela
         Call InstanciaClasse
         mvarMe.CmbIDTPMAQ.SetFocus
      Else
         mvarMe.CmbIDTPMAQ.SetFocus
      End If
   End If
End Sub
Private Sub mvarMe_CmdSairClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdSalvarClick()
   If Salvar Then
      If Not mvarShortBar Is Nothing Then
         mvarShortBar.TLCadastro.F_Refresh
      End If
   End If
End Sub
Private Sub mvarMe_Load()
   DoEvents
   Call MontaTela
   Call PopulaTela
   Screen.MousePointer = vbDefault
End Sub

