VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADMAQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe  As FrmCADMAQ
Attribute mvarMe.VB_VarHelpID = -1

Private mvarShortBar As TL_ShortBar
'Private WithEvents mvarMDI As FrmMDI

'Private mvarRptRec        As ReportRecordItem
Private mvarRptRec         As ReportRow
Private TbFunc             As Object 'TB_OMAQUINA

Private mvarIDLOJA         As String
Private mvarIDMAQUINA  As String
Private mvarSys            As Object
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
   mvarIDLOJA = RptRec.Record.Item(0).Value
   mvarIDMAQUINA = RptRec.Record.Item(1).Value
End Property
Public Property Get RptRec() As ReportRow
   Set RptRec = mvarRptRec
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmCADRFUNCIONARIO
'   Set mvarMDI = New FrmMDI
End Sub
Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call SetTag(mvarMe, "1VEZ", 0)
'      Call PopulaTela
   End If
End Sub
Private Sub InstanciaClasse()
   Set TbFunc = Nothing
   Set TbFunc = CriarObjeto("BANCO_3R.TB_RFUNCIONARIO")
   Set TbFunc.xDb = mvarSys.xDb
End Sub
Private Sub PopulaTela()
   If TbFunc Is Nothing Then
      Call InstanciaClasse
   End If
   
   With TbFunc
      If .pesquisar(Ch_IDMAQUINA:=mvarIDMAQUINA) Then
         mvarMe.CmdExcluir.Enabled = True
         
         mvarMe.TxtID.Text = .IDMAQUINA
         mvarMe.OptATIVO(IIf(.SITFUNC = "A", 1, 0)).Value = True
         Call LocalizarCombo(.CmbIDTPMAQ, .IDTPMAQ, True, False)
         mvarMe.TxtCODMAQUINA.Text = .CODMAQUINA
         mvarMe.TxtDTOPERACAO.Text = .DTOPERACAO
         
      Else
         Call LimpaTela
      End If
   End With
End Sub
Public Function Excluir() As Boolean
   If mvarIDMAQUINA <> 0 Then
      If VerificarDependencias() Then
         If ExibirPergunta("Confirma a exclusão da máquina.", "Atenção", False) = vbYes Then
            Call TbFunc.pesquisar(Ch_IDMAQUINA:=xVal(mvarMe.TxtID))
            If TbFunc.Excluir() Then
               Call ExibirResultado(mvarSys, True)
               
               'mvarIsDirt = True
               
               Call LimpaTela
               If mvarMe.TxtCODMAQUINA.Visible And mvarMe.TxtCODMAQUINA.Enabled Then
                  mvarMe.TxtCODMAQUINA.SetFocus
               End If
            Else
               Call ExibirResultado(mvarSys, False)
            End If
         End If
      End If
   End If
End Function
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
      .CmbIDTPMAQ.ListIndex = 0
   End With
End Sub
Private Sub MontaTela()
   
   Dim MyTb As Object
   Set MyTb = CriarObjeto("BANCO_3R.TB_OTPMAQ")
   Set MyTb.xDb = mvarSys.xDb
   If MyTb.Pesquisa Then
      With mvarMe.CmbIDTPMAQ
         .Clear
         While Not MyTb.Rs.EOF
            MyTb.Popula
            .AddItem MyTb.DSCMAQ, MyTb.IDTPMAQ
            MyTb.Rs.MoveNext
         Wend
      End With
   End If
End Sub
Private Sub PopulaClasse()
   If TbFunc Is Nothing Then Call InstanciaClasse
   
   With TbFunc
      If .IDMAQUINA = 0 Then .DTCADASTRO = Date
      .IDLOJA = mvarIDLOJA
      .CODMAQUINA = mvarMe.TxtCODMAQUINA.Text
      .DTOPERACAO = mvarMe.TxtDTOPERACAO.Text
      .IDTPMAQ = mvarMe.CmbIDTPMAQ.ItemData(mvarMe.CmbIDTPMAQ.ListIndex)
   End With
End Sub
Private Function Salvar() As Boolean
   Dim bOk     As Boolean
   Dim xNode   As TreeViewNode
   Dim bIsDirt As Boolean
   
   If ValidaCampos() Then
      Call PopulaClasse
      bIsDirt = TbFunc.isDirt
      If bIsDirt Then
         If TbFunc.Salvar() Then
            bOk = True
         End If
      Else
         bOk = True
      End If
      Call ExibirResultado(mvarSys, bOk, IIf(bOk, 1, 2))
   End If
   Salvar = bOk
End Function
Private Function VerificarDependencias()
   Dim bResult    As Boolean
   Dim TbAtend    As Object
   Dim TbVenda    As Object
   Dim sMsg       As String
   
   Set TbAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
   With TbAtend
      Set .xDb = mvarSys.xDb
      If .pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDMAQUINA=" & SqlNum(mvarMe.TxtID)) Then
         sMsg = "Funcionário já realizou atendimento e não pode ser excluído."
         Call ExibirAviso(sMsg)
         GoTo Saida
      Else
         Set TbVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
         With TbVenda
            Set .xDb = mvarSys.xDb
            If .pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDMAQUINA=" & SqlNum(mvarMe.TxtID)) Then
               sMsg = "Funcionário já realizou venda e não pode ser excluído."
               Call ExibirAviso(sMsg)
               GoTo Saida
            End If
         End With
      End If
   End With
   
   bResult = True
   
   VerificarDependencias = bResult
Saida:
   Set TbAtend = Nothing
   Set TbVenda = Nothing
End Function

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
         mvarMe.TxtCODMAQUINA.SetFocus
      Else
         mvarMe.TxtCODMAQUINA.SetFocus
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
End Sub

