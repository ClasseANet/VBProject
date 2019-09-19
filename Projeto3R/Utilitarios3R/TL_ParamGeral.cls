VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_ParamGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMDI As FrmMDI
Attribute mvarMDI.VB_VarHelpID = -1

Private WithEvents mvarMe  As FrmParamGeral
Attribute mvarMe.VB_VarHelpID = -1

Private mvarsys      As Object
Private mvarIDLOJA   As Integer
Public Property Set Sys(ByVal vData As Object)
    Set mvarsys = vData
    mvarIDLOJA = mvarsys.Propriedades("IDLOJA")
End Property
Public Property Get Sys() As Object
    Set Sys = mvarsys
End Property
Public Property Set MDI(ByVal vData As Object)
    Set mvarMDI = vData
End Property
Public Property Get MDI() As Object
    Set MDI = mvarMDI
End Property
Public Sub Show()
   '* Carregar e Acoplar Form no MDI
   Load mvarMe
   Call AcoplarForm(mvarMe, 2, mvarsys, pMDI:=mvarMDI)
End Sub
Private Sub mvarMDI_CmdCancelarClick()
   Call mvarMe_CmdCancelar
End Sub
Private Sub mvarMDI_CmdOkClick()
   Call mvarMe_CmdOk
End Sub
Private Sub mvarMDI_CmdPadraoClick()
   Call mvarMe_CmdPadrao
End Sub
Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 0) = 0 Then
      Call SetTag(mvarMe, "1VEZ", 1)
      If mvarMe.TxtFX1.Visible And mvarMe.TxtFX1.Enabled Then
         mvarMe.TxtFX1.SetFocus
      ElseIf mvarMe.ChkFaixaMeta.Visible And mvarMe.ChkFaixaMeta.Enabled Then
         mvarMe.ChkFaixaMeta.SetFocus
      End If
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_ChkFaixaMetaClick()
   With mvarMe
      .TxtFX1.Enabled = (.ChkFaixaMeta = xtpChecked)
      .TxtFX2.Enabled = (.ChkFaixaMeta = xtpChecked)
      .TxtFX3.Enabled = (.ChkFaixaMeta = xtpChecked)
      .ChkFXMetaQTD.Enabled = (.ChkFaixaMeta = xtpChecked)
   End With
End Sub
Private Sub mvarMe_ChkTPAREAClick()
   If mvarMe.ChkTPAREA.Value = xtpChecked Then
      mvarMe.GrpFoto.Enabled = True
      mvarMe.ChkTPDIR.Value = mvarsys.Getparam("TPDIR", Default:=0)
      mvarMe.ChkTPDISP.Value = mvarsys.Getparam("TPDISP", Default:=0)
   Else
      mvarMe.GrpFoto.Enabled = False
      mvarMe.ChkTPDIR.Value = 0
      mvarMe.ChkTPDISP.Value = 0
   End If
End Sub

Private Sub mvarMe_CmdCancelar()
   Call PopulaTela
End Sub
Private Sub mvarMe_CmdOk()
   Call SalvarTela
End Sub
Private Sub mvarMe_CmdPadrao()
   With mvarMe
      .TxtFX1.Text = 270
      .TxtFX2.Text = 320
      .TxtFX3.Text = 345
   End With
End Sub
Private Sub mvarMe_Load()
   Screen.MousePointer = vbHourglass
   'Call SetRunTimeFormProperty(mvarMe)
   Call MontaTela
   Call PopulaTela
End Sub
Private Sub mvarMe_Resize()
   With mvarMe
      .GrpTela.ZOrder 1
      .GrpBoxTop.Move 0, -90, .ScaleWidth - 60, 800
      .GrpTela.Move 0, 0, .ScaleWidth - 60, .ScaleHeight
      .GrpSenhaMestre.Move 180, 960
      .GrpCadastroUnico.Top = .GrpSenhaMestre.Top
      .GrpEstruturaServ.Top = .GrpCadastroUnico.Top + .GrpCadastroUnico.Height + 60
      .GrpIdentificacao.Top = .GrpSenhaMestre.Top + .GrpSenhaMestre.Height + 60
      .GrpIdentificacao.Left = .GrpSenhaMestre.Left
      .GrpSalas.Top = .GrpIdentificacao.Top
      .GrpSalas.Left = .GrpCadastroUnico.Left
      .GrpCalendario.Top = .GrpIdentificacao.Top + .GrpIdentificacao.Height + 60
      .GrpCalendario.Left = .GrpIdentificacao.Left
      
   End With
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmParamGeral
End Sub
Private Sub MontaTela()
   Dim Sql As String
   Dim MyRs As Object
   
   With mvarMe
      .ChkOCLIENTE1.Enabled = False
      .ChkOCONTATO1.Enabled = False
      .ChkRFUNCIONARIO1.Enabled = False
      
      Sql = "Select Count(*) [QTD] From OLOJA"
      If mvarsys.xdb.AbreTabela(Sql, MyRs) Then
         If xVal(mvarsys.xdb.RsAux("QTD")) > 1 Then
            .ChkOCLIENTE1.Enabled = True
            .ChkOCONTATO1.Enabled = True
            .ChkRFUNCIONARIO1.Enabled = True
         End If
      End If
      
      With .CmbDia1Semana
         .Clear
         .AddItem "Domingo", 0
         .ItemData(.NewIndex) = 1
         .AddItem "Segunda-Feira", 1
         .ItemData(.NewIndex) = 2
         .AddItem "Terça-Feira", 2
         .ItemData(.NewIndex) = 3
         .AddItem "Quarta-Feira", 3
         .ItemData(.NewIndex) = 4
         .AddItem "Quinta-Feira", 4
         .ItemData(.NewIndex) = 5
         .AddItem "Sexta-Feira", 5
         .ItemData(.NewIndex) = 6
         .AddItem "Sábado", 6
         .ItemData(.NewIndex) = 7
      End With
  End With
End Sub
Private Sub PopulaTela()
   Dim TbMeta  As Object
   Call LimpaTela

   Set TbMeta = CriarObjeto("Banco_3R.TB_PMETA")
   With TbMeta
      Set .xdb = mvarsys.xdb
      If Not .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDMETA:=1) Then
         .IDLOJA = mvarIDLOJA
         .IDMETA = 1
         .FAIXA1 = 270
         .FAIXA2 = 320
         .FAIXA3 = .FAIXA2 + ((.FAIXA2 - .FAIXA1) / 2)
         .Salvar
         Call .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDMETA:=1)
      End If
   End With
  With mvarMe
      .TxtFX1.Text = TbMeta.FAIXA1
      .TxtFX2.Text = TbMeta.FAIXA2
      .TxtFX3.Text = TbMeta.FAIXA3
      
      .ChkFaixaMeta.Value = mvarsys.Getparam("VIEWFXMETA", Default:=0)
      .ChkFXMetaQTD.Value = mvarsys.Getparam("VIEWFXMETAQTD", Default:=0)
      
      .ChkOCLIENTE1.Enabled = False
      .ChkOCONTATO1.Enabled = False
      .ChkRFUNCIONARIO1.Enabled = False
      If mvarsys.xdb.AbreTabela("Select Count(*) [QTD] From OLOJA") Then
         If xVal(mvarsys.xdb.RsAux("QTD")) > 0 Then
            .ChkOCLIENTE1.Enabled = True
            .ChkOCONTATO1.Enabled = True
            .ChkRFUNCIONARIO1.Enabled = True
            
            .ChkOCLIENTE1.Value = mvarsys.Getparam("OCLIENTEUNICO", Default:=0)
            .ChkOCONTATO1.Value = mvarsys.Getparam("OCONTATOUNICO", Default:=0)
            .ChkRFUNCIONARIO1.Value = mvarsys.Getparam("RFUNCIONARIOUNICO", Default:=0)
         End If
      End If
      
      .ChkTPSERV.Value = mvarsys.Getparam("TPSERV", Default:=1)
      .ChkTPTRAT.Value = mvarsys.Getparam("TPTRAT", Default:=1)
      .ChkTPAREA.Value = mvarsys.Getparam("TPAREA", Default:=1)
      
      .ChkTPDIR.Value = mvarsys.Getparam("TPDIR", Default:=0)
      .ChkTPDISP.Value = mvarsys.Getparam("TPDISP", Default:=0)
      
      .ChkBIOMETRIA.Value = mvarsys.Getparam("BIOMETRIA", Default:=0)
      .ChkPonto.Value = mvarsys.Getparam("PONTO", Default:=0)
      .ChkExibeSala.Value = mvarsys.Getparam("EXIBESALA", Default:=1)
      
      .TxtStartTime.Text = Format(mvarsys.Getparam("WorkDayStartTime", Default:="09:00"), "hh:mm")
      .TxtEndTime.Text = Format(mvarsys.Getparam("WorkDayEndTime", Default:="17:00"), "hh:mm")
  End With
End Sub
Private Sub SalvarTela()
   Dim TbParam As Object
   Dim TbMeta  As Object
   Dim Queries As Collection
   
   Set Queries = New Collection
   
   With mvarMe
      Set TbMeta = CriarObjeto("Banco_3R.TB_PMETA")
      Set TbMeta.xdb = mvarsys.xdb
      TbMeta.IDLOJA = mvarIDLOJA
      TbMeta.IDMETA = 1
      TbMeta.FAIXA1 = xVal(.TxtFX1.Text)
      TbMeta.FAIXA2 = xVal(.TxtFX2.Text)
      TbMeta.FAIXA3 = xVal(.TxtFX3.Text)
      Queries.Add TbMeta.QrySave
      
      Queries.Add QrySaveParam(mvarsys, "VIEWFXMETA", .ChkFaixaMeta.Value)
      Queries.Add QrySaveParam(mvarsys, "VIEWFXMETAQTD", .ChkFXMetaQTD.Value)
      
      Queries.Add QrySaveParam(mvarsys, "OCLIENTEUNICO", .ChkOCLIENTE1.Value)
      Queries.Add QrySaveParam(mvarsys, "OCONTATOUNICO", .ChkOCONTATO1.Value)
      Queries.Add QrySaveParam(mvarsys, "RFUNCIONARIOUNICO", .ChkRFUNCIONARIO1.Value)
            
      Queries.Add QrySaveParam(mvarsys, "TPSERV", .ChkTPSERV.Value)
      Queries.Add QrySaveParam(mvarsys, "TPTRAT", .ChkTPTRAT.Value)
      Queries.Add QrySaveParam(mvarsys, "TPAREA", .ChkTPAREA.Value)
      
      Queries.Add QrySaveParam(mvarsys, "TPDIR", .ChkTPDIR.Value)
      Queries.Add QrySaveParam(mvarsys, "TPDISP", .ChkTPDISP.Value)
      
      Queries.Add QrySaveParam(mvarsys, "BIOMETRIA", .ChkBIOMETRIA.Value)
      Queries.Add QrySaveParam(mvarsys, "PONTO", .ChkPonto.Value)
      Queries.Add QrySaveParam(mvarsys, "EXIBESALA", .ChkExibeSala.Value)
   
      Queries.Add QrySaveParam(mvarsys, "WorkDayStartTime", Format(.TxtStartTime.Text, "hh:mm"))
      Queries.Add QrySaveParam(mvarsys, "WorkDayEndTime", Format(.TxtEndTime.Text, "hh:mm"))
                     
      If mvarsys.xdb.executa(Queries) Then
         Dim sAux As String
         
         Call ExibirResultado(mvarsys, True, 1)
         mvarsys.Propriedades("VIEWFXMETA") = .ChkFaixaMeta.Value
         mvarsys.Propriedades("VIEWFXMETAQTD") = .ChkFXMetaQTD.Value
         mvarsys.Propriedades("BIOMETRIA") = .ChkBIOMETRIA.Value
         mvarsys.Propriedades("PONTO") = .ChkPonto.Value
         
         mvarsys.Propriedades("TPSERV") = .ChkTPSERV.Value
         mvarsys.Propriedades("TPTRAT") = .ChkTPTRAT.Value
         mvarsys.Propriedades("TPAREA") = .ChkTPAREA.Value
         
         mvarsys.Propriedades("TPDIR") = .ChkTPDIR.Value
         mvarsys.Propriedades("TPDISP") = .ChkTPDISP.Value
      Else
         Call ExibirResultado(mvarsys, False, 1)
      End If
   End With
End Sub
Private Sub LimpaTela()
   Dim i As Integer
   On Error Resume Next
   With mvarMe
      For i = 0 To .Controls.Count - 1
         .Controls(i).Text = ""
         .Controls(i).Value = 0
      Next
   End With
End Sub
Private Sub mvarMe_TxtEndTimeLostFocus()
   mvarMe.TxtEndTime.Text = FormatarHora(mvarMe.TxtEndTime.Text)
End Sub
Private Sub mvarMe_TxtStartTimeLostFocus()
   mvarMe.TxtStartTime.Text = FormatarHora(mvarMe.TxtStartTime.Text)
End Sub
