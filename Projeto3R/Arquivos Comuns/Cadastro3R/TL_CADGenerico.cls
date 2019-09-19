VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADGenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe  As FrmCADGenerico
Attribute mvarMe.VB_VarHelpID = -1

Private mvarShortBar As TL_ShortBar
'Private WithEvents mvarMDI As FrmMDI

Private mvarRptRec         As ReportRow

Private mvarTbName   As String
Private mvarTbObj    As Object
Private mvarID       As String
Private mvarSys      As Object
Public Property Set ShortBar(ByVal vData As TL_ShortBar)
   Set mvarShortBar = vData
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Property Let Id(ByVal vData As String)
   mvarID = vData
End Property
Public Property Get Id() As String
   Id = mvarID
End Property
Public Property Let TbName(ByVal vData As String)
   mvarTbName = vData
End Property
Public Property Get TbName() As String
   TbName = mvarTbName
End Property
Public Property Set TBObj(ByVal vData As Object)
   Set mvarTbObj = vData
End Property
Public Property Get TBObj() As Object
   Set TBObj = mvarTbObj
End Property
Public Property Set RptRec(ByVal vData As ReportRow)
   Set mvarRptRec = vData
   If Not mvarRptRec Is Nothing Then mvarID = mvarRptRec.Record.Item(0).Value
   'mvarArayCampoIDFUNCIONARIO = RptRec.Record.Item(1).Value
End Property
Public Property Get RptRec() As ReportRow
   Set RptRec = mvarRptRec
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmCADGenerico
'   Set mvarMDI = New FrmMDI
End Sub
Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call SetTag(mvarMe, "1VEZ", 0)
'      Call PopulaTela
   End If
End Sub
Private Sub InstanciaClasse()
   Set TBObj = Nothing
   Set TBObj = CriarObjeto("BANCO_3R." & mvarTbName)
   Set TBObj.Xdb = mvarSys.Xdb
End Sub
Private Sub PopulaTela()
'   If TbObj Is Nothing Then
'      Call InstanciaClasse
'   End If
'
'   With TbObj
'      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDFUNCIONARIO:=mvarID) Then
'         mvarMe.CmdExcluir.Enabled = True
'
'         mvarMe.TxtID.Text = .IDFUNCIONARIO
'         mvarMe.OptATIVO(IIf(.SITFUNC = "A", 1, 0)).Value = True
'         mvarMe.txtNome.Text = .NOME
'         mvarMe.TxtCHAPA.Text = .CHAPA
'         mvarMe.TxtDTADMISSAO.Text = .DTADMISSAO
'         mvarMe.TxtDTDEMISSAO.Text = .DTDEMISSAO
'         mvarMe.TxtDTNASC.Text = .DTNASC
'         mvarMe.ChkFLGCERTIFICADO.Value = .FLGCERTIFICADO
'         mvarMe.TxtObs.Text = .OBS
'
'         'Call LocalizarCombo(mvarMe.CmbSexo, .SEXO, True, False)
'         mvarMe.txtEndereco.Text = .ENDERECO
'         mvarMe.txtBairro.Text = .BAIRRO
'         mvarMe.txtCidade.Text = .CIDADE
'         mvarMe.TxtEstado.Text = .ESTADO
'         mvarMe.cmbPais.Text = .PAIS
'         mvarMe.txtCEP.Text = .CEP
'         mvarMe.txtCelular.Text = .CELULAR
'         mvarMe.txtTel.Text = .TELEFONE
'         mvarMe.txtEmail.Text = .EMAIL
'
'         mvarMe.TxtSALARIO.Text = .SALARIO
'         mvarMe.ChkCOMPROD.Value = .COMPROD
'         mvarMe.ChkCOMSERV.Value = .COMSERV
'         mvarMe.TxtVLCOMPROD.Text = .VLCOMPROD
'         mvarMe.TxtVLCOMSERV.Text = .VLCOMSERV
'         mvarMe.OptTPCOMPROD(.TPCOMPROD).Value = True
'         mvarMe.OptTPCOMSERV(.TPCOMSERV).Value = True
'      Else
'         Call LimpaTela
'      End If
'   End With
End Sub
Public Function Excluir() As Boolean
   If mvarID <> 0 Then
'      If VerificarDependencias() Then
'         If ExibirPergunta("Confirma a exclusão do funcionário.", "Atenção", False) = vbYes Then
'            Call TbObj.Pesquisar(Ch_IDLOJA:=mvarIDLOJA,Ch_IDFUNCIONARIO:=xVal(mvarMe.TxtID), Ch_IDLOJA:=mvarIDLOJA)
'            If TbObj.Excluir() Then
'               Call ExibirResultado(mvarSys, True)
'
'               'mvarIsDirt = True
'
'               Call LimpaTela
'               If mvarMe.txtNome.Visible And mvarMe.txtNome.Enabled Then
'                  mvarMe.txtNome.SetFocus
'               End If
'            Else
'               Call ExibirResultado(mvarSys, False)
'            End If
'         End If
'      End If
   End If
End Function
Private Function ValidaCampos() As Boolean
   ' *******************
   Dim bResult As Boolean
   
   bResult = True
   If Trim(mvarMe.TxtCampo01.Text) = "" Then
      ExibirAviso "O campo deve ser preenchido."
      bResult = False
      mvarMe.TxtCampo01.SetFocus
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
   End With
End Sub
Private Sub MontaTela()

End Sub
Private Sub PopulaClasse()
   If TBObj Is Nothing Then Call InstanciaClasse
   
'   With TBObj
'      If .IDFUNCIONARIO = 0 Then .DTCADASTRO = Date
'      .IDLOJA = mvarIDLOJA
'      .NOME = mvarMe.txtNome.Text
'      .CHAPA = mvarMe.TxtCHAPA.Text
'      .SITFUNC = IIf(mvarMe.OptATIVO(0).Value, "I", "A")
'      .FLGCERTIFICADO = mvarMe.ChkFLGCERTIFICADO.Value
'      .DTADMISSAO = mvarMe.TxtDTADMISSAO.Text
'      .DTDEMISSAO = mvarMe.TxtDTDEMISSAO.Text
'      .DTNASC = mvarMe.TxtDTNASC.Text
'      .OBS = mvarMe.TxtObs.Text
'
'      .ENDERECO = mvarMe.txtEndereco.Text
'      .BAIRRO = mvarMe.txtBairro.Text
'      .CIDADE = mvarMe.txtCidade.Text
'      .ESTADO = mvarMe.TxtEstado.Text
'      .PAIS = mvarMe.cmbPais.Text
'      .CEP = mvarMe.txtCEP.Text
'
'      .TELEFONE = mvarMe.txtCelular.Text
'      .EMAIL = mvarMe.txtEmail.Text
'      .CELULAR = mvarMe.txtCelular.Text
'
'      .SALARIO = xVal(mvarMe.TxtSALARIO.Text)
'      .COMPROD = mvarMe.ChkCOMPROD.Value
'      .COMSERV = mvarMe.ChkCOMSERV.Value
'      .VLCOMPROD = xVal(mvarMe.TxtVLCOMPROD.Text)
'      .VLCOMSERV = xVal(mvarMe.TxtVLCOMSERV.Text)
'      .TPCOMPROD = IIf(mvarMe.OptTPCOMPROD(0).Value, 0, 1)
'      .TPCOMSERV = IIf(mvarMe.OptTPCOMSERV(0).Value, 0, 1)
'
'   End With
End Sub
Private Function Salvar() As Boolean
   Dim bOk     As Boolean
   Dim xNode   As TreeViewNode
   Dim bIsDirt As Boolean
   
   If ValidaCampos() Then
      Call PopulaClasse
      bIsDirt = TBObj.isDirt
      If bIsDirt Then
         If TBObj.Salvar() Then
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
   
'   Set TbAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
'   With TbAtend
'      Set .XDb = mvarSys.XDb
'      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDFUNCIONARIO=" & SqlNum(mvarMe.txtID)) Then
'         sMsg = "Funcionário já realizou atendimento e não pode ser excluído."
'         Call ExibirAviso(sMsg)
'         GoTo Saida
'      Else
'         Set TbVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
'         With TbVenda
'            Set .XDb = mvarSys.XDb
'            If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDFUNCIONARIO=" & SqlNum(mvarMe.txtID)) Then
'               sMsg = "Funcionário já realizou venda e não pode ser excluído."
'               Call ExibirAviso(sMsg)
'               GoTo Saida
'            End If
'         End With
'      End If
'   End With
   
   bResult = True
   
   VerificarDependencias = bResult
Saida:
   Set TbAtend = Nothing
   Set TbVenda = Nothing
End Function
Private Sub mvarMe_CmdExcluirClick()
   If mvarID <> 0 Then
      If mvarMe.Visible Then
         Call PopulaTela
      End If
      Call Excluir
   End If
End Sub
Private Sub mvarMe_CmdNovoClick()
   If xVal(mvarMe.TxtID) = 0 Then
      Call LimpaTela
      mvarMe.TxtCampo01.SetFocus
   Else
      If Salvar Then
         Call LimpaTela
         Call InstanciaClasse
         mvarMe.TxtCampo01.SetFocus
      Else
         mvarMe.TxtCampo01.SetFocus
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
   Call MontaTela
   Call PopulaTela
End Sub

