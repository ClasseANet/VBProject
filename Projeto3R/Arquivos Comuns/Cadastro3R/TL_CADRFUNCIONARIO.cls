VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADRFUNCIONARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe  As FrmCADRFUNCIONARIO
Attribute mvarMe.VB_VarHelpID = -1

Private mvarShortBar As TL_ShortBar
'Private WithEvents mvarMDI As FrmMDI

'Private mvarRptRec        As ReportRecordItem
Private mvarRptRec         As ReportRow
Private TbFunc             As Object 'TB_RFUNCIONARIO

Private mvarIDLOJA         As String
Private mvarIDFUNCIONARIO  As String

Private mvarSys            As Object
Public Property Set TLCadastro(ByRef vData As Object)
   Dim sAux As String
   With vData
      .Titulo = "Funcionários"
      
      sAux = "Select IDLOJA, IDFUNCIONARIO, CHAPA, NOME"
      sAux = sAux & ", DTADMISSAO [ADMISSÃO], DTDEMISSAO [DEMISSÃO]"
      sAux = sAux & " From RFUNCIONARIO"
      sAux = sAux & " Where IDLOJA=" & mvarIDLOJA
      .Consulta = sAux
      
      sAux = ""
      sAux = sAux & "|CAMPO=IDLOJA       |CHAVE=S|GRUPO=0|VISIBLE=N:"
      sAux = sAux & "|CAMPO=IDFUNCIONARIO|CHAVE=S|GRUPO=0|VISIBLE=N:"
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
Public Property Let IDFUNCIONARIO(ByVal vData As String)
   mvarIDFUNCIONARIO = vData
End Property
Public Property Get IDFUNCIONARIO() As String
   IDFUNCIONARIO = mvarIDFUNCIONARIO
End Property
Public Property Set RptRec(ByVal vData As ReportRow)
   Set mvarRptRec = vData
   If Not mvarRptRec Is Nothing Then mvarIDFUNCIONARIO = mvarRptRec.Record.Item(1).Value
End Property
Public Property Get RptRec() As ReportRow
   Set RptRec = mvarRptRec
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmCADRFUNCIONARIO
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
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDFUNCIONARIO:=mvarIDFUNCIONARIO) Then
         mvarMe.CmdExcluir.Enabled = True
         
         mvarMe.TxtID.Text = .IDFUNCIONARIO
         mvarMe.OptATIVO(IIf(.SITFUNC = "A", 1, 0)).Value = True
         mvarMe.TxtNOME.Text = .NOME
         Call LocalizarCombo(mvarMe.CmbIDLOJA0, .IDLOJA0, True, True)
         If mvarMe.CmbIDLOJA0.ListIndex = -1 Then
            If mvarMe.CmbIDLOJA0.ListCount = 1 Then
               mvarMe.CmbIDLOJA0.ListIndex = 0
            End If
         End If
         mvarMe.TxtCHAPA.Text = .CHAPA
         mvarMe.TxtDTADMISSAO.Text = .DTADMISSAO
         mvarMe.TxtDTDEMISSAO.Text = .DTDEMISSAO
         mvarMe.TxtDTNASC.Text = .DTNASC
         mvarMe.ChkFLGCERTIFICADO.Value = .FLGCERTIFICADO
         mvarMe.ChkFLGVENDA.Value = .FLGVENDA
         mvarMe.TxtSenha.Text = Decrypt2(.SENHA)
         mvarMe.TxtObs.Text = .OBS
         
         'Call LocalizarCombo(mvarMe.CmbSexo, .SEXO, True, False)
         mvarMe.TxtEndereco.Text = .ENDERECO
         mvarMe.TxtBairro.Text = .BAIRRO
         mvarMe.TxtCidade.Text = .CIDADE
         mvarMe.TxtESTADO.Text = .ESTADO
         mvarMe.cmbPais.Text = .PAIS
         mvarMe.TxtCEP.Text = .CEP
         mvarMe.txtCelular.Text = .CELULAR
         mvarMe.txtTel.Text = .TELEFONE
         mvarMe.TxtEMAIL.Text = .EMAIL
         
         mvarMe.TxtSALARIO.Text = .SALARIO
         mvarMe.ChkCOMPROD.Value = .COMPROD
         mvarMe.ChkCOMSERV.Value = .COMSERV
         mvarMe.TxtVLCOMPROD.Text = .VLCOMPROD
         mvarMe.TxtVLCOMSERV.Text = .VLCOMSERV
         mvarMe.OptTPCOMPROD(.TPCOMPROD).Value = True
         mvarMe.OptTPCOMSERV(.TPCOMSERV).Value = True
         
         Call LocalizarCombo(mvarMe.CmbDIAFOLGA, .DIAFOLGA, True, True)
      Else
         Call LimpaTela
         If mvarMe.CmbIDLOJA0.ListIndex = -1 Then
            If mvarMe.CmbIDLOJA0.ListCount = 1 Then
               mvarMe.CmbIDLOJA0.ListIndex = 0
            Else
               Call LocalizarCombo(mvarMe.CmbIDLOJA0, mvarIDLOJA, True, True)
            End If
         End If
      End If
   End With
End Sub
Public Function Excluir() As Boolean
   If mvarIDFUNCIONARIO <> 0 Then
      If VerificarDependencias() Then
         If ExibirPergunta("Confirma a exclusão do funcionário.", "Atenção", False) = vbYes Then
            Call TbFunc.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDFUNCIONARIO:=xVal(mvarMe.TxtID))
            If TbFunc.Excluir() Then
               Call ExibirResultado(mvarSys, True)
               
               'mvarIsDirt = True
               
               Call LimpaTela
               If Not mvarShortBar Is Nothing Then
                  mvarShortBar.TLCadastro.F_Refresh
               End If
               If mvarMe.TxtNOME.Visible And mvarMe.TxtNOME.Enabled Then
                  mvarMe.TxtNOME.SetFocus
               End If
               Excluir = True
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
'   If Trim(mvarMe.txtTel.Text) = "" And Trim(mvarMe.txtCelular.Text) = "" And Trim(mvarMe.txtEmail.Text) = "" Then
'      mvarMe.TabContato.Item(1).Selected = True
'      ExibirAviso "Ao menos um contato, telefone ou e-Mail, deve ser preenchido."
'      bResult = False
'      mvarMe.txtCelular.SetFocus
'      GoTo Fim
'   End If
   
   If Trim(mvarMe.TxtNOME.Text) = "" Then
      mvarMe.TabContato.Item(0).Selected = True
      ExibirAviso "O nome deve ser preenchido."
      bResult = False
      mvarMe.TxtNOME.SetFocus
      GoTo Fim
   End If
   If Trim(mvarMe.TxtCHAPA.Text) = "" Then
      mvarMe.TabContato.Item(0).Selected = True
      'ExibirAviso "Informe o registro do funcionário." & vbNewLine & "O Sistema irá sugerir um código."
      bResult = False
      If mvarSys.xDb.AbreTabela("Select Max(CHAPA) [COD] From RFUNCIONARIO Where IDLOJA=" & mvarIDLOJA) Then
         mvarMe.TxtCHAPA.Text = StrZero(xVal(mvarSys.xDb.RsAux("COD") & "") + 1, 6)
      Else
         mvarMe.TxtCHAPA.Text = StrZero(1, 6)
      End If
      
      'mvarMe.TxtCHAPA.SetFocus
      'GoTo Fim
   ElseIf Len(Trim(mvarMe.TxtCHAPA.Text)) <> 6 Then
      mvarMe.TxtCHAPA.Text = StrZero(Trim(mvarMe.TxtCHAPA.Text), 6)
   End If
   
   mvarMe.TxtDTADMISSAO.Text = FormatarData(mvarMe.TxtDTADMISSAO.Text)
   If Trim(mvarMe.TxtDTADMISSAO.Text) = "" Then
      mvarMe.TabContato.Item(0).Selected = True
      ExibirAviso "Admissão inválida."
      bResult = False
      mvarMe.TxtDTADMISSAO.SetFocus
      GoTo Fim
   End If
   
   If xVal(mvarMe.TxtID.Text) = 0 Then
      If mvarMe.ChkFLGCERTIFICADO.Value = xtpUnchecked Then
         If vbYes = ExibirPergunta("Funcionário é certificado a operar equipamentos?") Then
            mvarMe.ChkFLGCERTIFICADO.Value = xtpChecked
         End If
      End If
      If mvarMe.ChkFLGVENDA.Value = xtpUnchecked Then
         If vbYes = ExibirPergunta("Funcionário realizará vendas?") Then
            mvarMe.ChkFLGVENDA.Value = xtpChecked
         End If
      End If
      bResult = True
   End If
   
Fim:
   ValidaCampos = bResult
End Function
Public Sub Limpar()
   mvarIDFUNCIONARIO = ""
End Sub
Private Sub LimpaTela()
   Dim n As Object
   On Error Resume Next
   With mvarMe
      For Each n In mvarMe.Controls
         n.Text = ""
      Next
      .OptATIVO(1).Value = True
      .ChkFLGCERTIFICADO.Value = xtpUnchecked
      .ChkFLGVENDA.Value = xtpUnchecked
      .OptTPCOMPROD(0).Value = True
      .OptTPCOMSERV(0).Value = True
      .ChkCOMPROD.Value = xtpUnchecked
      .ChkCOMSERV.Value = xtpUnchecked
   End With
End Sub
Private Sub MontaTela()
   mvarMe.TabContato.Item(0).Selected = True
   mvarMe.CmdBiometria.Enabled = (xVal(mvarSys.Propriedades("BIOMETRIA")) = 1)
   mvarMe.CmdBiometria.Visible = (xVal(mvarSys.Propriedades("BIOMETRIA")) = 1)
   
   With mvarMe.CmbDIAFOLGA
      .Clear
      .AddItem " "
      .ItemData(.NewIndex) = 0
      .AddItem "02 - Segunda-Feira"
      .ItemData(.NewIndex) = 2
      .AddItem "03 - Terça-Feira"
      .ItemData(.NewIndex) = 3
      .AddItem "04 - Quarta-Feira"
      .ItemData(.NewIndex) = 4
      .AddItem "05 - Quinta-Feira"
      .ItemData(.NewIndex) = 5
      .AddItem "06 - Sexta-Feira"
      .ItemData(.NewIndex) = 6
      .AddItem "07 - Sábado"
      .ItemData(.NewIndex) = 7
   End With
   
   Dim TbLoja As Object
   Set TbLoja = CriarObjeto("BANCO_3R.TB_OLOJA")
   Set TbLoja.xDb = mvarSys.xDb
   mvarMe.CmbIDLOJA0.Clear
   If TbLoja.Pesquisar(Ch_WHERE:="IDCOLIGADA=" & mvarSys.IDCOLIGADA) Then
      While Not TbLoja.Rs.EOF
         TbLoja.Popula
         mvarMe.CmbIDLOJA0.AddItem TbLoja.NOME
         mvarMe.CmbIDLOJA0.ItemData(mvarMe.CmbIDLOJA0.NewIndex) = TbLoja.IDLOJA
         TbLoja.Rs.MoveNext
      Wend
   End If
   If mvarMe.CmbIDLOJA0.ListCount = 1 Then
      mvarMe.CmbIDLOJA0.ListIndex = 0
      mvarMe.CmbIDLOJA0.Enabled = False
   End If
      
   
   
   
End Sub
Private Sub PopulaClasse()
   If TbFunc Is Nothing Then Call InstanciaClasse
   
   With TbFunc
      If .IDFUNCIONARIO = 0 Then .DTCADASTRO = Date
      .IDCOLIGADA = mvarSys.IDCOLIGADA
      .IDLOJA = mvarIDLOJA
      .IDLOJA0 = mvarMe.CmbIDLOJA0.ItemData(mvarMe.CmbIDLOJA0.ListIndex)
      .NOME = mvarMe.TxtNOME.Text
      .CHAPA = mvarMe.TxtCHAPA.Text
      .SITFUNC = IIf(mvarMe.OptATIVO(0).Value, "I", "A")
      .FLGCERTIFICADO = mvarMe.ChkFLGCERTIFICADO.Value
      .FLGVENDA = mvarMe.ChkFLGVENDA.Value
      .DTADMISSAO = mvarMe.TxtDTADMISSAO.Text
      .DTDEMISSAO = mvarMe.TxtDTDEMISSAO.Text
      .DTNASC = mvarMe.TxtDTNASC.Text
      .SENHA = Encrypt2(mvarMe.TxtSenha.Text)
      .OBS = mvarMe.TxtObs.Text
            
      .ENDERECO = mvarMe.TxtEndereco.Text
      .BAIRRO = mvarMe.TxtBairro.Text
      .CIDADE = mvarMe.TxtCidade.Text
      .ESTADO = mvarMe.TxtESTADO.Text
      .PAIS = mvarMe.cmbPais.Text
      .CEP = mvarMe.TxtCEP.Text
      
      .TELEFONE = mvarMe.txtCelular.Text
      .EMAIL = mvarMe.TxtEMAIL.Text
      .CELULAR = mvarMe.txtCelular.Text
      
      .SALARIO = xVal(mvarMe.TxtSALARIO.Text)
      .COMPROD = mvarMe.ChkCOMPROD.Value
      .COMSERV = mvarMe.ChkCOMSERV.Value
      .VLCOMPROD = xVal(mvarMe.TxtVLCOMPROD.Text)
      .VLCOMSERV = xVal(mvarMe.TxtVLCOMSERV.Text)
      .TPCOMPROD = IIf(mvarMe.OptTPCOMPROD(0).Value, 0, 1)
      .TPCOMSERV = IIf(mvarMe.OptTPCOMSERV(0).Value, 0, 1)
      
      .DIAFOLGA = mvarMe.CmbDIAFOLGA.ItemData(mvarMe.CmbDIAFOLGA.ListIndex)
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
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_WHERE:="IDFUNCIONARIO=" & SqlNum(mvarMe.TxtID)) Then
         sMsg = "Funcionário já realizou atendimento e não pode ser excluído."
         Call ExibirAviso(sMsg)
         GoTo Saida
      Else
         Set TbVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
         With TbVenda
            Set .xDb = mvarSys.xDb
            If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_WHERE:="IDFUNCIONARIO=" & SqlNum(mvarMe.TxtID)) Then
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

Private Sub mvarMe_ChkCOMPRODClick()
   With mvarMe
      .LblVLCOMPROD.Enabled = (.ChkCOMPROD.Value = xtpChecked)
      .TxtVLCOMPROD.Enabled = (.ChkCOMPROD.Value = xtpChecked)
      .OptTPCOMPROD(0).Enabled = (.ChkCOMPROD.Value = xtpChecked)
      .OptTPCOMPROD(1).Enabled = (.ChkCOMPROD.Value = xtpChecked)
   
      If .ChkCOMPROD.Value = xtpChecked Then
         If .TxtVLCOMPROD.Visible And .TxtVLCOMPROD.Enabled Then
            .TxtVLCOMPROD.SetFocus
         End If
      End If
   End With
End Sub

Private Sub mvarMe_ChkCOMSERVClick()
   With mvarMe
      .LblVLCOMSERV.Enabled = (.ChkCOMSERV.Value = xtpChecked)
      .TxtVLCOMSERV.Enabled = (.ChkCOMSERV.Value = xtpChecked)
      .OptTPCOMSERV(0).Enabled = (.ChkCOMSERV.Value = xtpChecked)
      .OptTPCOMSERV(1).Enabled = (.ChkCOMSERV.Value = xtpChecked)
      
      If .ChkCOMSERV.Value = xtpChecked Then
         If .TxtVLCOMSERV.Visible And .TxtVLCOMSERV.Enabled Then
            .TxtVLCOMSERV.SetFocus
         End If
      End If
   End With

End Sub

Private Sub mvarMe_CmdBiometriaClick()
   Dim ClFinger As Object
   Dim Sql As String
   
   Set ClFinger = CriarObjeto("FingerTech.NG_Bio")
   
   With ClFinger
      .Biometria = (xVal(mvarSys.Propriedades("BIOMETRIA")) = 1)
      If .Biometria Then
         Call .Init_Finger
         .FDbPath = Environ("PROGRAMFILES") & "\ClasseA\Projeto3R\Digitais\"
         .FDbName = mvarSys.CODSIS & ".fdb"
         
         If .RegisterFinger(xVal(mvarIDFUNCIONARIO)) Then
            Sql = "Update RFUNCIONARIO"
            Sql = Sql & " Set IDFINGER=IDFUNCIONARIO "
            Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
            Sql = Sql & " And IDFUNCIONARIO=" & mvarIDFUNCIONARIO
            If mvarSys.xDb.Executa(Sql) Then
               Call ExibirAviso("Biometria salva com sucesso!")
            Else
               Call ExibirStop("Erro ao salvar biometria")
            End If
         Else
         End If
         
         .Terminate_Finger
      End If
   End With
   Set ClFinger = Nothing
End Sub

Private Sub mvarMe_CmdExcluirClick()
   If mvarIDFUNCIONARIO <> 0 Then
      If mvarMe.Visible Then
         Call PopulaTela
      End If
      Call Excluir
   End If
End Sub
Private Sub mvarMe_CmdNovoClick()
   If xVal(mvarMe.TxtID) = 0 And Trim(mvarMe.TxtNOME) = "" Then
      Call LimpaTela
      mvarMe.TxtNOME.SetFocus
   Else
      If Salvar Then
         Call LimpaTela
         Call InstanciaClasse
         mvarMe.TxtNOME.SetFocus
      Else
         mvarMe.TxtNOME.SetFocus
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
