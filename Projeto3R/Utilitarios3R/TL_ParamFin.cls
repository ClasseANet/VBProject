VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_ParamFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMDI As FrmMDI
Attribute mvarMDI.VB_VarHelpID = -1

Private WithEvents mvarMe  As FrmParamFin
Attribute mvarMe.VB_VarHelpID = -1

Private mvarSys      As Object
Private mvarIDLOJA   As Integer
Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
    mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
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
   Call AcoplarForm(mvarMe, 2, mvarSys, pMDI:=mvarMDI)
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
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmdCancelar()
   Call PopulaTela
End Sub
Private Sub mvarMe_CmdOk()
   Call SalvarTela
End Sub
Private Sub mvarMe_CmdPadrao()
   With mvarMe
      .ChkNOMEFAVORECIDO.Value = xtpChecked
   End With
End Sub
Private Sub mvarMe_LblUltFatDblClick()
   Dim sLim As String
   If SenhaMestre(mvarSys, 1) Then
      sLim = InputBox("Limite Faturamento", "Parâmetros", mvarSys.Getparam("NFEX", Default:=0, pSECAO:="FIN", pDescricao:="Param NF"))
      If Trim(sLim) <> "" Then
         Call mvarSys.SaveParam("NFEX", xVal(sLim))
      End If
   End If
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
      .GrpExibe.Move 180, .GrpBoxTop.Height, .ScaleWidth - 360
      .GrpIntegra.Move 180, .GrpExibe.Top + .GrpExibe.Height + 60, .GrpExibe.Width
      .GrpTaxas.Move 180, .GrpIntegra.Top + .GrpIntegra.Height + 60, .GrpIntegra.Width
   End With
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmParamFin
End Sub
Private Sub MontaTela()
   Dim Sql As String
   Dim MyRs As Object
   
   With mvarMe
      .ChkNOMEFAVORECIDO.Enabled = True
      .ChkNFE.Enabled = True
      .ChKNFE_CLI.Enabled = True
      .ChkNFE_CPF.Enabled = True
   End With
End Sub
Private Sub PopulaTela()
   Dim oRs As Object
   Dim Sql As String
   
   On Error Resume Next
   With mvarMe
      .ChkNOMEFAVORECIDO.Value = mvarSys.Getparam("FINNMFAVORECIDO", Default:=0, pSECAO:="FIN", pDescricao:="Exibição de Cliente em Favorecido")
      .ChkNFE.Value = mvarSys.Getparam("NFE", Default:=0, pSECAO:="FIN", pDescricao:="Integração NF-e")
      .ChKNFE_CLI.Value = mvarSys.Getparam("NFECLI", Default:=0, pSECAO:="FIN", pDescricao:="Enviar Recibo Cliente")
      .ChkNFE_CPF.Value = mvarSys.Getparam("NFECPF", Default:=0, pSECAO:="FIN", pDescricao:="Recibo por CPF")
      
      .ChkNFE.Enabled = (.ChkNFE.Value = xtpChecked)
      .ChKNFE_CLI.Enabled = (.ChkNFE.Value = xtpChecked)
      .ChkNFE_CPF.Enabled = (.ChkNFE.Value = xtpChecked)
      .TxtTXSERV2.Text = ""
      .TxtTXSERV3.Text = ""
      
      Sql = "Select IDFORMAPGTO, TXSERV, TXPARC "
      Sql = Sql & " From CFORMAPGTO"
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And IDFORMAPGTO in (2,3)"
      If mvarSys.xDb.AbreTabela(Sql, oRs) Then
         While Not oRs.EOF
            If oRs("IDFORMAPGTO") = 2 Then
               .TxtTXSERV2.Text = ValBr(xVal(oRs("TXSERV") & ""))
            ElseIf oRs("IDFORMAPGTO") = 3 Then
               .TxtTXSERV3.Text = ValBr(xVal(oRs("TXSERV") & ""))
               .TxtTXSERV4.Text = ValBr(xVal(oRs("TXPARC") & ""))
            End If
            oRs.MoveNext
         Wend
      End If
      
      .TxtUltFat.Text = mvarSys.Getparam("ULTFAT", Default:=0, pSECAO:="FIN", pDescricao:="Último Mês de Faturamento Enviado")
   End With
End Sub
Private Sub SalvarTela()
   Dim TbParam As Object
   Dim TbMeta  As Object
   Dim Queries As Collection
      
   Set Queries = New Collection
   
   With mvarMe
      Call mvarSys.SaveParam("FINNMFAVORECIDO", .ChkNOMEFAVORECIDO.Value)
      Call mvarSys.SaveParam("NFE", .ChkNFE.Value)
      Call mvarSys.SaveParam("NFECLI", .ChKNFE_CLI.Value)
      Call mvarSys.SaveParam("NFECPF", .ChkNFE_CPF.Value)
      Call mvarSys.SaveParam("ULTFAT", xVal(.TxtUltFat.Text), pCATPARAM:="FIN", pDSCPARAM:="Último Mês de Faturamento Enviado")
      
      Queries.Add "Update CFORMAPGTO Set TXSERV=" & Replace(xVal(mvarMe.TxtTXSERV2.Text), ",", ".") & ", ALTERSTAMP=1, TIMESTAMP=GETDATE() Where IDFORMAPGTO=2"
      Queries.Add "Update CFORMAPGTO Set TXSERV=" & Replace(xVal(mvarMe.TxtTXSERV3.Text), ",", ".") & ", ALTERSTAMP=1, TIMESTAMP=GETDATE() Where IDFORMAPGTO=3"
      Queries.Add "Update CFORMAPGTO Set TXPARC=" & Replace(xVal(mvarMe.TxtTXSERV4.Text), ",", ".") & ", ALTERSTAMP=1, TIMESTAMP=GETDATE() Where IDFORMAPGTO=3"
      If mvarSys.xDb.Executa(Queries) Then
         Call ExibirResultado(mvarSys, True, 1)
      '   mvarSys.Propriedades("VIEWFXMETA") = .ChkFaixaMeta.Value
      Else
         Call ExibirResultado(mvarSys, False, 1)
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

