VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADOTPTRATAMENTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const TbName = "OTPTRATAMENTO"
Const IdName = "IDTPTRATAMENTO"
Const Campo1 = "DSCTRATAMENTO"

Private WithEvents mvarMe  As FrmCADOTPTRATAMENTO
Attribute mvarMe.VB_VarHelpID = -1

Private mvarShortBar       As TL_ShortBar
Private mvarRptRec         As ReportRow
Private TbCad              As Object
Private mvarSys            As Object
Private mvarIDLOJA         As Integer
Private mvarIDTPTRATAMENTO As Integer
Private mvarDSCTRAT        As String
Public Property Set TLCadastro(ByRef vData As Object)
   Dim Sql  As String
   Dim sAux As String
   With vData
      .Titulo = "Tratamentos"
      
      Sql = "Select " & IdName & " [Id.]"
      Sql = Sql & ", " & Campo1 & " [Descrição]"
      Sql = Sql & ", FREQUENCIA [Frequência] "
      Sql = Sql & ", FLGDISPARO [Equipamento] "
      Sql = Sql & ", FLGAREA [Área] "
      Sql = Sql & ", FLGAVALIACAO [Avaliação] "
      Sql = Sql & ", ATIVO [Ativo] "
      
      Sql = Sql & " From " & TbName
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      Sql = Sql & " AND FLGDEL<>'X'"
      'SQL = SQL & " And IDTPTRATAMENTO<>0"
      .Consulta = Sql
      
      sAux = ""
      sAux = sAux & "|CAMPO=" & IdName & "|CHAVE=S|GRUPO=0|VISIBLE=S:"
      .TAGCAMPOS = sAux
   End With
End Property
Public Property Let DSCTRAT(ByVal vData As String)
   mvarDSCTRAT = vData
End Property
Public Property Get DSCTRAT() As String
   DSCTRAT = mvarDSCTRAT
End Property
Public Property Let IDTPTRATAMENTO(ByVal vData As String)
   mvarIDTPTRATAMENTO = vData
End Property
Public Property Get IDTPTRATAMENTO() As String
   IDTPTRATAMENTO = mvarIDTPTRATAMENTO
End Property
Public Property Set RptRec(ByVal vData As ReportRow)
   Set mvarRptRec = vData
'   mvarIDLOJA = RptRec.Record.Item(0).Value
   mvarIDTPTRATAMENTO = 0
   If Not mvarRptRec Is Nothing Then
      mvarIDTPTRATAMENTO = mvarRptRec.Record.Item(0).Value
   End If
End Property
Public Property Get RptRec() As ReportRow
   Set RptRec = mvarRptRec
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
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmCADOTPTRATAMENTO
End Sub
Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call SetTag(mvarMe, "1VEZ", 0)
      If mvarDSCTRAT <> "" And mvarMe.TxtID.Text = "" Then
         mvarMe.TxtCampo01.Text = mvarDSCTRAT
      End If
      If mvarMe.TxtCampo01.Enabled And mvarMe.TxtCampo01.Visible Then
         mvarMe.TxtCampo01.SetFocus
      End If
      
      'Call MontaTela
      'Call PopulaTela
   End If
End Sub
Private Sub InstanciaClasse()
   Set TbCad = Nothing
   Set TbCad = CriarObjeto("BANCO_3R.TB_" & TbName, False)
   Set TbCad.xDb = mvarSys.xDb
End Sub
Private Sub PopulaTela()
   Dim Sql As String
   Dim bExiste As Boolean
   
   If TbCad Is Nothing Then
      Call InstanciaClasse
   End If
   
   With TbCad
      bExiste = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDTPTRATAMENTO:=mvarIDTPTRATAMENTO)
      'If bExiste And mvarIDTPTRATAMENTO <> 0 Then
      If bExiste Then
      
         If xVal(.COR) = 0 Then
            .COR = CorPadrao(.DSCTRATAMENTO)
            If .isDirt Then .Salvar
         End If
         mvarMe.CmdExcluir.Enabled = True
         
         mvarMe.TxtID.Text = .IDTPTRATAMENTO
         mvarMe.TxtCampo01.Text = .DSCTRATAMENTO
         mvarMe.TxtCampo02.Text = .FREQUENCIA
         mvarMe.ChkFLGDISPARO.Value = .FLGDISPARO
         mvarMe.ChkFLGAREA.Value = .FLGAREA
         mvarMe.ChkFLGAVALIACAO.Value = .FLGAVALIACAO
         mvarMe.ChkATIVO.Value = .ATIVO
         mvarMe.CtlColor.Color = .COR
         
         Call PopulaProd
         
      Else
         Call LimpaTela
         If mvarDSCTRAT <> "" Then
            mvarMe.TxtCampo01.Text = mvarDSCTRAT
            If mvarMe.TxtCampo01.Enabled And mvarMe.TxtCampo01.Visible Then
               mvarMe.TxtCampo01.SetFocus
            End If
         End If
      End If
      mvarMe.CmdPadrao.Enabled = (CorPadrao(mvarMe.TxtCampo01.Text) <> RGB(255, 255, 255))
   End With
End Sub
Private Function PopulaProd() As Boolean
   Dim MyRs As Object
   Dim bOk As Boolean
   Dim Sql As String
   
'   If mvarMe.ChkProdVenda.Value = xtpUnchecked Then
'      mvarMe.ChkProdVenda.Value = xtpChecked
'      Exit Function
'   End If
      
   Sql = "Select S.IDPROD, S.CODPROD, S.NMPROD, S.VLVENDA"
   Sql = Sql & " From OTRATAMENTO_PROD T"
   Sql = Sql & " Join SPRODUTO S On S.IDLOJA=T.IDLOJA And S.IDPROD=T.IDPROD"
   Sql = Sql & " Where T.IDLOJA=" & SqlNum(mvarIDLOJA)
   Sql = Sql & " And T.IDTPTRATAMENTO=" & SqlNum(mvarMe.TxtID.Text)
   bOk = mvarSys.xDb.AbreTabela(Sql, MyRs)
   If Not bOk Then
      Sql = "Select S.IDPROD, S.CODPROD, S.NMPROD, S.VLVENDA"
      Sql = Sql & " From SPRODUTO S "
      Sql = Sql & " Where S.IDLOJA=" & SqlNum(mvarIDLOJA)
      Sql = Sql & " And S.NMPROD = " & SqlStr(mvarMe.TxtCampo01.Text)
      bOk = mvarSys.xDb.AbreTabela(Sql, MyRs)
   End If
   
   If bOk Then
      PopulaProd = True
      mvarMe.ChkProdVenda.Value = xtpChecked
      mvarMe.TxtCodProd.Tag = MyRs("CODPROD")
      mvarMe.TxtCodProd.Text = MyRs("CODPROD")
      mvarMe.TxtNMPROD.Text = MyRs("NMPROD")
      mvarMe.TxtVLVENDA.Text = ValBr(MyRs("VLVENDA"))
   Else
      PopulaProd = False
      mvarMe.TxtCodProd.Tag = ""
      mvarMe.TxtCodProd.Text = ""
      If mvarMe.ChkProdVenda.Value = xtpChecked Then
         mvarMe.TxtNMPROD.Text = mvarMe.TxtCampo01.Text
      Else
         mvarMe.TxtNMPROD.Text = ""
      End If
      mvarMe.TxtVLVENDA.Text = ""
   End If
End Function
Private Function SalvarProd() As Boolean

   Dim pQueries As Collection
   Dim TbProd As Object
   Dim TbTratProd As Object
   Dim bOk As Boolean
   
   Set pQueries = New Collection
   
   If Trim(mvarMe.TxtNMPROD.Text) <> "" And mvarMe.ChkProdVenda.Value = xtpChecked Then
      Set TbProd = Nothing
      Set TbProd = CriarObjeto("BANCO_3R.TB_SPRODUTO", False)
      Set TbProd.xDb = mvarSys.xDb
      Call TbProd.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPROD:=xVal(mvarMe.TxtCodProd.Tag))
      TbProd.IDLOJA = mvarIDLOJA
      TbProd.CODPROD = mvarMe.TxtCodProd.Text
      TbProd.NMPROD = mvarMe.TxtNMPROD.Text
      TbProd.VLVENDA = xVal(mvarMe.TxtVLVENDA.Text)
      
      TbProd.DTCADASTRO = mvarSys.xDb.SysDate()
      TbProd.DSCPROD = TbProd.NMPROD
      TbProd.UNIDCONTROLE = "un"
      TbProd.UNIDCOMPRA = TbProd.UNIDCONTROLE
      TbProd.UNIDVENDA = TbProd.UNIDCONTROLE
      TbProd.VLMEDIO = TbProd.VLVENDA
      TbProd.EVENDA = 1
      TbProd.ESERVICO = 1
      TbProd.SITPROD = "A"
      TbProd.FLGDEL = "0"
      TbProd.IDUSU = mvarSys.IDUSU
      'pQueries.Add TbProd.QrySave
      If TbProd.Salvar Then
      
         Set TbTratProd = Nothing
         Set TbTratProd = CriarObjeto("BANCO_3R.TB_OTRATAMENTO_PROD", False)
         Set TbTratProd.xDb = mvarSys.xDb
         Call TbTratProd.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDTPTRATAMENTO=" & mvarIDTPTRATAMENTO)
         TbTratProd.IDLOJA = mvarIDLOJA
         TbTratProd.IDPROD = TbProd.IDPROD
         TbTratProd.IDTPSERVICO = 2
         TbTratProd.IDTPTRATAMENTO = mvarIDTPTRATAMENTO
         TbTratProd.IDAREA = 0
         
         If TbTratProd.Salvar Then
            bOk = True
         End If
      End If
   End If
   If bOk Then
      mvarMe.TxtCodProd.Tag = TbProd.IDPROD
      mvarMe.TxtCodProd.Text = TbProd.CODPROD
      mvarMe.TxtNMPROD.Text = TbProd.NMPROD
      mvarMe.TxtVLVENDA.Text = ValBr(TbProd.VLVENDA)
   Else
      If mvarMe.ChkProdVenda.Value = xtpChecked Then
         Call ExibirStop("Produto de Venda NÃO foi salvo!")
      End If
      mvarMe.TxtCodProd.Tag = ""
      mvarMe.TxtCodProd.Text = ""
      mvarMe.TxtVLVENDA.Text = ""
   End If
   
End Function
Public Function CorPadrao(pDSCTRATAMENTO As String) As Long
   Dim nCor As Long
   Select Case UCase(pDSCTRATAMENTO)
      Case "DEPILAÇÃO", "DEPILACAO":   nCor = RGB(165, 222, 99)
      Case "REJUVENESCIMENTO":         nCor = RGB(166, 222, 255) 'RGB(165, 206, 198)
      Case "ACNE", "MANCHA":           nCor = RGB(214, 206, 132)
      Case "NARL":                     nCor = RGB(250, 210, 95)
      Case "LIMPEZA DE PELE":          nCor = RGB(255, 132, 255)
      Case "PEELING":                  nCor = RGB(249, 203, 254) '(247, 180, 254)
      Case Else
         If InStr(UCase(pDSCTRATAMENTO), "FACIAL") <> 0 Then
            nCor = RGB(255, 132, 255)
         ElseIf InStr(UCase(pDSCTRATAMENTO), "PEELING") <> 0 Then
            nCor = RGB(249, 203, 254)
         ElseIf InStr(UCase(pDSCTRATAMENTO), "MASSAGEM") <> 0 Then
            nCor = RGB(253, 251, 172)
         ElseIf InStr(UCase(pDSCTRATAMENTO), "SOBRANCELHA") <> 0 Then
            nCor = RGB(178, 178, 178)
         Else
            nCor = RGB(255, 255, 255)
         End If
   End Select
   CorPadrao = nCor
End Function
Public Function Excluir() As Boolean
   Dim Queries As Collection
   Dim bDep As Boolean
   Dim TbProd As Object
   Dim Sql As String
   
   If mvarIDTPTRATAMENTO >= 0 Then
      bDep = VerificarDependencias
'      If VerificarDependencias() Then
         If ExibirPergunta("Esta operação irá excluir o tratamento." & vbNewLine & vbNewLine & "Confirma?", "Atenção", False) = vbYes Then
            Call TbCad.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDTPTRATAMENTO:=xVal(mvarMe.TxtID))
            Set Queries = New Collection
            If bDep Then
               TbCad.FLGDEL = "X"
               TbCad.ATIVO = 0
               Queries.Add TbCad.QryUpdate
               
               Set TbProd = Nothing
               Set TbProd = CriarObjeto("BANCO_3R.TB_SPRODUTO", False)
               Set TbProd.xDb = mvarSys.xDb
               If TbProd.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPROD:=xVal(mvarMe.TxtCodProd.Tag)) Then
                  TbProd.FLGDEL = "X"
                  TbProd.SITPROD = "I"
                  Queries.Add TbProd.QryUpdate
               End If
            Else
               Queries.Add TbCad.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDTPTRATAMENTO:=xVal(mvarMe.TxtID))
            End If
            If mvarSys.xDb.Executa(Queries) Then
               Call ExibirResultado(mvarSys, True)
               
               'mvarIsDirt = True
               
               Call LimpaTela
               If Not mvarShortBar Is Nothing Then
                  mvarShortBar.TLCadastro.F_Refresh
                  Sql = "Select Min(IDTPTRATAMENTO) [ID] From OTPTRATAMENTO Where IDLOJA=" & mvarIDLOJA & " And IDTPTRATAMENTO>" & mvarIDTPTRATAMENTO & " And FLGDEL<>'X'"
                  If mvarSys.xDb.AbreTabela(Sql) Then
                     mvarIDTPTRATAMENTO = mvarSys.xDb.RsAux("ID")
                     Call PopulaTela
                  End If
               End If
               If mvarMe.TxtID.Visible And mvarMe.TxtCampo01.Enabled Then
                  mvarMe.TxtCampo01.SetFocus
               End If
               Excluir = True
            Else
               Call ExibirResultado(mvarSys, False)
            End If
            Set Queries = Nothing
         End If
'      End If
   End If
End Function
Private Function ValidaCampos() As Boolean
   ' *******************
   Dim bResult As Boolean
   
   bResult = True
   If Trim(mvarMe.TxtCampo01.Text) = "" Then
      ExibirAviso "Descrição inválida."
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
      For Each n In mvarMe.Controls: n.Text = ""
      Next
      '.ChkFLGDISPARO.Value = xtpUnchecked
      '.ChkFLGAREA.Value = xtpUnchecked
      '.ChkFLGAVALIACAO.Value = xtpUnchecked
      mvarMe.TxtCodProd.Tag = ""
      .ChkATIVO.Value = xtpChecked
      .CtlColor.Color = RGB(255, 255, 255) '16777215  ' RGB(255, 255, 255)
   End With
End Sub
Private Sub MontaTela()
   Dim MyTb As Object
   mvarMe.TxtCampo01.MaxLength = 30
   'mvarMe.ChkFLGDISPARO.Caption = "Com Controle de Máquina"
End Sub
Private Sub PopulaClasse(ByRef Queries As Collection)
   Dim oTbSLMaq As Object
   Dim bExiste As Boolean
   
'   If TbCad Is Nothing Then
   Call InstanciaClasse
'   End If

   With TbCad
      bExiste = False
      If mvarMe.TxtID.Text <> "" Then
         bExiste = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDTPTRATAMENTO:=mvarMe.TxtID.Text)
      End If
'      If .IDTPTRATAMENTO = 0 Then
'         .DTCADASTRO = mvarSys.Xdb.SysDate()
'      End If
      .IDLOJA = mvarIDLOJA
      .DSCTRATAMENTO = mvarMe.TxtCampo01.Text
      .FREQUENCIA = xVal(mvarMe.TxtCampo02.Text)
      .FLGDISPARO = mvarMe.ChkFLGDISPARO.Value
      .FLGAREA = mvarMe.ChkFLGAREA.Value
      .FLGAVALIACAO = mvarMe.ChkFLGAVALIACAO.Value
      .ATIVO = mvarMe.ChkATIVO.Value
      .COR = mvarMe.CtlColor.Color
      .FLGDEL = "0"
      'If .isDirt Or Not bExiste Then
      '   Queries.Add .QrySave
      'End If
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
      
      bIsDirt = TbCad.isDirt
      If bIsDirt Then
         If TbCad.Salvar() Then
            bOk = True
            mvarIDTPTRATAMENTO = TbCad.IDTPTRATAMENTO
            mvarMe.TxtID.Text = mvarIDTPTRATAMENTO
            Call SalvarProd
         End If
      Else
         bOk = True
         Call SalvarProd
      End If
      Call ExibirResultado(mvarSys, bOk, IIf(bOk, 1, 2))
   End If
   Salvar = bOk
End Function
Private Function VerificarDependencias()
   Dim bResult    As Boolean
   Dim TbServ     As Object
   Dim sMsg       As String
   
   Set TbServ = CriarObjeto("BANCO_3R.TB_OSERVICOEVT")
   With TbServ
      Set .xDb = mvarSys.xDb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & SqlNum(mvarMe.TxtID)) Then
         'sMsg = "Este item já possui compromisso associado e não pode ser excluído."
         'sMsg = sMsg & vbNewLine & vbNewLine & "Deseja continuar?"
         bResult = True
         'Call ExibirAviso(sMsg)
         'bResult = (ExibirPergunta(sMsg, "Tratamento", False) = vbYes)
      Else
         bResult = False
      End If
   End With
   VerificarDependencias = bResult
Saida:
   Set TbServ = Nothing
End Function

Private Sub mvarMe_ChkProdVendaClick()
   If mvarMe.ChkProdVenda.Value = xtpChecked Then
      mvarMe.LblVLVENDA.ForeColor = vbBlack
      mvarMe.TxtCodProd.ForeColor = vbBlack
      mvarMe.TxtNMPROD.ForeColor = vbBlack
      mvarMe.TxtVLVENDA.Enabled = True
      mvarMe.TxtCodProd.Enabled = True
      mvarMe.TxtNMPROD.Enabled = True
      If Not PopulaProd Then
         mvarMe.TxtNMPROD.Text = mvarMe.TxtCampo01.Text
      End If
   Else
      mvarMe.LblVLVENDA.ForeColor = mvarMe.Label5.ForeColor
      mvarMe.TxtCodProd.ForeColor = mvarMe.Label5.ForeColor
      mvarMe.TxtNMPROD.ForeColor = mvarMe.Label5.ForeColor
      mvarMe.TxtVLVENDA.Enabled = False
      mvarMe.TxtCodProd.Tag = ""
      mvarMe.TxtCodProd.Text = ""
      mvarMe.TxtNMPROD.Text = ""
      mvarMe.TxtVLVENDA.Text = ""
   End If
End Sub

Private Sub mvarMe_CmdExcluirClick()
'   If mvarIDTPTRATAMENTO <> 0 Then
      If mvarMe.Visible Then
         Call PopulaTela
      End If
      Call Excluir
'   End If
End Sub
Private Sub mvarMe_CmdNovoClick()
   If xVal(mvarMe.TxtID) = 0 And Trim(mvarMe.TxtCampo01) = "" Then
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
Private Sub mvarMe_CmdPadraoClick()
   mvarMe.CtlColor.Color = CorPadrao(mvarMe.TxtCampo01.Text)
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
   Call LimpaTela
   Call MontaTela
   Call PopulaTela
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarMe_TxtCampo01Change()
   If mvarMe.TxtCodProd.Tag = "" And mvarMe.ChkProdVenda.Value = xtpChecked Then
      mvarMe.TxtNMPROD.Text = mvarMe.TxtCampo01.Text
   End If
End Sub

Private Sub mvarMe_TxtCODPRODLostFocus()
   Dim Sql As String
   Dim MyRs As Object
   
   If xVal(mvarMe.TxtCodProd.Text) <> xVal(mvarMe.TxtCodProd.Tag) Then
      Sql = "Select S.IDPROD, S.CODPROD, S.NMPROD, S.VLVENDA"
      Sql = Sql & " From SPRODUTO S "
      Sql = Sql & " Where S.IDLOJA=" & SqlNum(mvarIDLOJA)
      Sql = Sql & " And S.IDPROD = " & SqlNum(mvarMe.TxtCodProd.Text)
      If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
         mvarMe.ChkProdVenda.Value = xtpChecked
         mvarMe.TxtCodProd.Tag = MyRs("CODPROD")
         mvarMe.TxtCodProd.Text = MyRs("CODPROD")
         mvarMe.TxtNMPROD.Text = MyRs("NMPROD")
         mvarMe.TxtVLVENDA.Text = ValBr(MyRs("VLVENDA"))
      Else
         Call PopulaProd
      End If
   End If
End Sub

Private Sub mvarMe_TxtIDLostFocus()
   mvarIDTPTRATAMENTO = xVal(mvarMe.TxtID.Text)
   If Not mvarMe.ActiveControl Is mvarMe.CmdSair Then
      Call PopulaTela
   End If
   'mvarMe.
End Sub

