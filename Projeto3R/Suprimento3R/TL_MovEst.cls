VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_MovEst"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const StrPesquisa = "Pesquisar Estoque"
Private WithEvents mvarMe     As FrmMovEst
Attribute mvarMe.VB_VarHelpID = -1
Private WithEvents mvarPane   As FrmPaneSupri
Attribute mvarPane.VB_VarHelpID = -1


Private RsLan As Recordset  'Object 'Object
Dim iCampo As Collection
Private bNovo As Boolean

Const TaskCC_Mov = 1
Const TaskCC_Det = 2
Const TaskCC_His = 3

Private mvarSys      As Object
Private mvarIDLOJA   As Integer
Private mvarIDEST    As Integer
Private mvarIDPROD   As Long
Private mvarIDMOV    As Long
Private mvarUNIDCONTROLE As String
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
   mvarIDEST = mvarSys.Propriedades("IDEST")
   mvarIDPROD = ReadIniFile(mvarSys.LocalReg, "Config", "IDPROD", "2")
End Property
Public Property Get Sys() As Object
   Set Sys = mvarSys
End Property
Public Property Set Pane(ByVal vData As Object)
   Set mvarPane = vData
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDEST(ByVal vData As Integer)
   mvarIDEST = vData
End Property
Public Property Get IDEST() As Integer
   IDEST = mvarIDEST
End Property
Public Property Let IDPROD(ByVal vData As Long)
   mvarIDPROD = vData
End Property
Public Property Get IDPROD() As Long
   IDPROD = mvarIDPROD
End Property
Public Property Get UNIDCONTROLE() As String
   UNIDCONTROLE = mvarUNIDCONTROLE
End Property
Public Property Get IDMOV() As Long
   IDMOV = mvarIDMOV
End Property
Public Sub Show(Optional Reload As Boolean = False)
   If mvarIDEST = 0 Then
      Call ExibirAviso("Não existe estoque definido.")
   End If
   
   If Reload Then
      Set Sys = Sys
      Unload mvarMe
   End If
   
   '* Carregar e Acoplar Form no MDI
   Load mvarMe
   Call AcoplarForm(mvarMe, 2, mvarSys)
   mvarMe.Show
   
   '* Definir foco
   With mvarMe.GrdMov
      If .Visible And .Enabled Then
         .SetFocus
      End If
   End With
   If Reload Then Call mvarMe_Rezise
End Sub
Public Sub Hide()
   mvarMe.Hide
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmMovEst
   mvarIDPROD = 2
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_Activate()
   'Call PopulaTela
   If GetTag(mvarMe, "LOAD", 1) = 1 Then
      Call SetTag(mvarMe, "LOAD", 0)
      Call mvarMe_TabContaSelectedChanged(mvarMe.TabConta.Item(mvarMe.TabConta.SelectedItem))
      Call mvarMe_GrdMovSelectionChanged
      Call mvarMe_Rezise
   End If
   Screen.MousePointer = vbDefault
End Sub
'Private Sub mvarMe_CmbDeLostFocus()
'   Call PopulaUltTransf
'End Sub
Private Sub mvarMe_CmbDTMOVChange()
   mvarMe.TxtDTMOV.Text = Format(mvarMe.CmbDTMOV.Value, "dd/mm/yyyy")
End Sub

Private Sub mvarMe_CmbFavorecidoLostFocus()
   If bNovo Then
      If LocalizarCombo(mvarMe.CmbFavorecido, mvarMe.CmbFavorecido.Text, False) >= 0 Then
         Call LocalizarCombo(mvarMe.CmbFavorecido, mvarMe.CmbFavorecido.Text)
         'Call PopulaUltLan(mvarMe.CmbFavorecido.ItemData(mvarMe.CmbFavorecido.ListIndex))
      End If
   End If
End Sub
'Private Sub PopulaUltTransf()
'   Dim Sql        As String
'   Dim oRs        As Object
'   Dim bAchou     As Boolean
'   Dim sData      As String
'   Dim nIDEST   As Long
'
'   If mvarMe.CmbPara.ListIndex < 0 Then Exit Sub
'   If mvarMe.TabConta.SelectedItem <> 1 Then Exit Sub
'   If xVal(mvarMe.TxtValor.Text) <> 0 And Trim(mvarMe.CmbFavorecido.Text) <> "" Then Exit Sub
'
'   sData = Format(mvarMe.TxtDTMOV.Text, "dd/mm/yyyy")
'   If Not IsDate(sData) Then
'      sData = Format(mvarSys.xDb.sysdate, "dd/mm/yyyy")
'   End If
'   sData = sData & " 23:59"
'   nIDEST = mvarMe.CmbPara.ItemData(mvarMe.CmbPara.ListIndex)
'
'   Sql = "Select Top 1 IDFOR, QTDITEM, OBS, Max(DTMOV) [Data]"
'   Sql = Sql & " From FLAN"
'   Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
'   Sql = Sql & " And IDEST=" & SqlNum(nIDEST)
'   Sql = Sql & " And TPLAN='C'"
'   Sql = Sql & " AND TPTRANSA='T'"
'   Sql = Sql & " And DTMOV <= " & SqlStr(sData)
'   Sql = Sql & " Group By IDFOR, QTDITEM, OBS"
'   Sql = Sql & " Order By Data Desc"
'
'   bAchou = mvarSys.xDb.AbreTabela(Sql, oRs)
'
'   If Not bAchou Then
'      Sql = "Select Top 1 IDFOR, QTDITEM, Min(DTMOV) [Data]"
'      Sql = Sql & " From SMOVEST"
'      Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
'      Sql = Sql & " And IDEST=" & SqlNum(nIDEST)
'      Sql = Sql & " And IDPROD=" & SqlNum(nIDEST)
'      Sql = Sql & " And TPLAN='C'"
'      Sql = Sql & " AND TPTRANSA='T'"
'      Sql = Sql & " And DTMOV > " & SqlStr(sData)
'      Sql = Sql & " Group By IDFOR, QTDITEM, OBS"
'      Sql = Sql & " Order By Data"
'
'      bAchou = mvarSys.xDb.AbreTabela(Sql, oRs)
'   End If
'   If bAchou Then
'      If Trim(mvarMe.CmbFavorecido.Text) = "" And xVal(oRs("IDFOR") & "") > 0 Then
'         mvarMe.CmbFavorecido.ListIndex = -1
'         Call LocalizarCombo(mvarMe.CmbFavorecido, oRs("IDFOR") & "", True, True)
'      End If
'      If xVal(mvarMe.TxtValor.Text) = 0 Then
'         mvarMe.TxtValor.Text = ValBr(xVal(oRs("VALOR") & ""),0)
'      End If
'      'If Trim(mvarMe.TxtObs.Text) <> "" Then
'      '   mvarMe.TxtObs.Text = oRs("OBS") & ""
'      'End If
'   End If
'End Sub
'Private Sub PopulaUltLan(pIDCONTATO As Long)
'   Dim Sql As String
'   Dim oRs As Object
'   Dim bAchou As Boolean
'   Dim sData  As String
'
'   If pIDCONTATO <= 0 Then
'      Exit Sub
'   End If
'
'   sData = Format(mvarMe.TxtDTMOV.Text, "dd/mm/yyyy")
'   If Not IsDate(sData) Then
'      sData = Format(mvarSys.xDb.sysdate, "dd/mm/yyyy")
'   End If
'   sData = sData & " 23:59"
'
'   Sql = "Select Top 1 QTDITEM, OBS, Max(DTMOV) [Data]"
'   Sql = Sql & " From FLAN"
'   Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
'   Sql = Sql & " And IDEST=" & SqlNum(mvarIDEST)
'   Sql = Sql & " And IDFOR=" & SqlNum(pIDCONTATO)
'   Sql = Sql & " And DTMOV <=" & SqlStr(sData)
'   Sql = Sql & " Group By IDDESP, IDSUBDESP, QTDITEM, OBS"
'   Sql = Sql & " Order By Data Desc"
'
'   bAchou = mvarSys.xDb.AbreTabela(Sql, oRs)
'   If Not bAchou Then
'      Sql = "Select Top 1 QTDITEM, OBS, Min(DTMOV) [Data]"
'      Sql = Sql & " From FLAN"
'      Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
'      Sql = Sql & " And IDEST=" & SqlNum(mvarIDEST)
'      Sql = Sql & " And IDFOR=" & SqlNum(pIDCONTATO)
'      Sql = Sql & " And DTMOV >" & SqlStr(sData)
'      Sql = Sql & " Group By IDDESP, IDSUBDESP, QTDITEM, OBS"
'      Sql = Sql & " Order By Data"
'
'      bAchou = mvarSys.xDb.AbreTabela(Sql, oRs)
'   End If
'   If bAchou Then
'      If xVal(mvarMe.TxtValor.Text) = 0 Then
'         mvarMe.TxtValor.Text = ValBr(xVal(oRs("QTDITEM") & ""),0)
'      End If
'      If Trim(mvarMe.TxtObs.Text) <> "" Then
'         mvarMe.TxtObs.Text = oRs("OBS") & ""
'      End If
'   End If
'End Sub
Private Sub mvarMe_CmbParaLostFocus()
'   Call PopulaUltTransf
End Sub
Private Sub mvarMe_CmdCancelarClick()
   Call HabilitaTab(False)
   Call PopulaTab
   mvarMe.GrdMov.SetFocus
End Sub
Private Sub mvarMe_CmdNovoClick()
   Dim sData   As String
   Dim nIndex As Long
   
   
   If mvarMe.GrdMov.SelectedRows.Count > 0 Then
      nIndex = mvarMe.GrdMov.SelectedRows(0).Index
      sData = Format(mvarMe.GrdMov.SelectedRows(0).Record.Item(iCampo("DTMOV")).Value, "dd/mm/yyyy")
   End If
   Call HabilitaTab(True)
   Call InserirGrdLinha
      
   If mvarMe.GrdMov.Rows.Count > nIndex Then
      mvarMe.GrdMov.Rows(nIndex).Selected = True
   End If
   If IsDate(sData) Then
      mvarMe.TxtDTMOV.Text = sData
      mvarMe.CmbDTMOV.Value = CDate(mvarMe.TxtDTMOV.Text)
   End If
      
   If mvarMe.TabConta.SelectedItem = 1 Then
      'Call SelContaTrans
   Else
      If mvarMe.CmbFavorecido.Visible And mvarMe.CmbFavorecido.Enabled Then
         mvarMe.CmbFavorecido.SetFocus
      End If
      'If mvarMe.TxtNUMDOC.Visible And mvarMe.TxtNUMDOC.Enabled Then
      '   mvarMe.TxtNUMDOC.SetFocus
      'End If
   
   End If
   bNovo = True
End Sub
'Private Sub SelContaTrans()
'   DoEvents
'   mvarMe.CmbDe.ListIndex = -1
'   mvarMe.CmbDe.ListIndex = LocalizarCombo(mvarMe.CmbDe, CStr(mvarIDEST), False, True)
'   If mvarMe.CmbDe.ListIndex >= 0 Then
'      If mvarMe.CmbPara.ListCount = 2 Then
'         mvarMe.CmbPara.ListIndex = 0
'         If mvarMe.CmbPara.ItemData(0) = mvarMe.CmbDe.ItemData(mvarMe.CmbDe.ListIndex) Then
'            mvarMe.CmbPara.ListIndex = 1
'         End If
'         If mvarMe.TxtDTMOV.Enabled And mvarMe.TxtDTMOV.Visible Then
'            mvarMe.TxtDTMOV.SetFocus
'         End If
'      Else
'         If mvarMe.CmbPara.Enabled And mvarMe.CmbPara.Visible Then
'            mvarMe.CmbPara.SetFocus
'         End If
'      End If
'   Else
'      If mvarMe.CmbDe.Enabled And mvarMe.CmbDe.Visible Then
'         mvarMe.CmbDe.SetFocus
'      End If
'   End If
'End Sub
Private Sub mvarMe_CmdNumClick()
   Dim NgFinanc As Object
   
   Set NgFinanc = CriarObjeto("Financ3R.NG_Financeiro")
   Set NgFinanc.Sys = mvarSys
   Call NgFinanc.ExibirVenda(mvarMe.TxtNUMDOC.Text)
End Sub
Private Sub mvarMe_CmdSalvarClick()
   If mvarMe.ActiveControl Is mvarMe.TxtDTMOV Then
      Call mvarMe_TxtDTMOVLostFocus
   End If
   Call SalvarLAN
End Sub
Private Sub SalvarLAN()
   Dim bOk        As Boolean
   Dim bIsDirt    As Boolean
   Dim TbSMOVEST     As Object
   Dim TbEstP     As Object 'TB_SESTPROD
   Dim TbTransf   As Object
   Dim nIndex     As Long
   Dim nTopRow    As Long
   Dim Queries    As Collection
   Dim nCtPara    As Integer
   Dim nCtDe      As Integer
   Dim oRptRow    As ReportRow
   Dim TbMovAntes As Object
   
   If Not ValidaCampos() Then Exit Sub
   If Not VerificaLAN(False, mvarIDMOV) Then Exit Sub
   
   Set Queries = New Collection
   
   nTopRow = mvarMe.GrdMov.TopRowIndex
   
   Set TbMovAntes = CriarObjeto("BANCO_3R.TB_SMOVEST")
   TbMovAntes.xDb = mvarSys.xDb
   Call TbMovAntes.Pesquisar(Ch_IDMOV:=mvarIDMOV)
   
   Call PopulaClasse(TbSMOVEST, Queries)
   
   With TbSMOVEST
      bIsDirt = .isDirt
      bOk = True
      If bIsDirt Then
         
         Set Queries = New Collection
         
         Set TbEstP = CriarObjeto("BANCO_3R.TB_SESTPROD")
         Set TbEstP.xDb = mvarSys.xDb
         
         Call TbEstP.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEST:=mvarIDEST, Ch_IDPROD:=.IDPROD)
         TbEstP.IDLOJA = mvarIDLOJA
         TbEstP.IDEST = mvarIDEST
         TbEstP.IDPROD = .IDPROD
         If TbEstP.UNIDCONTROLE = "" And .UNIDCONTROLE <> "" Then TbEstP.UNIDCONTROLE = .UNIDCONTROLE
         If .UNIDCONTROLE = "" And TbEstP.UNIDCONTROLE <> "" Then .UNIDCONTROLE = TbEstP.UNIDCONTROLE
         If TbEstP.UNIDCONTROLE = "" Then TbEstP.UNIDCONTROLE = "pç"
         If .UNIDCONTROLE = "" Then .UNIDCONTROLE = "pç"
         
         If TbEstP.UNIDCONTROLE <> .UNIDCONTROLE Then
            If ExibirPergunta("Unidade de medida está diferente do Estoque." & vbNewLine & vbNewLine & "Deseja contiuar?", pDefaultYes:=False) = vbNo Then
               GoTo Saida
            End If
         End If
         TbEstP.SLDATUAL = TbEstP.SLDATUAL - TbMovAntes.QTDITEM + xVal(.QTDITEM)
         TbEstP.SLDDISPONIVEL = TbEstP.SLDDISPONIVEL - TbMovAntes.QTDITEM + xVal(.QTDITEM)
         
         If TbEstP.SLDATUAL < 0 Or TbEstP.SLDDISPONIVEL < 0 Then
            If ExibirPergunta("Saldo de produto negativo." & vbNewLine & vbNewLine & "Deseja contiuar?", , pDefaultYes:=False) = vbNo Then
               GoTo Saida
            End If
         End If
            
         Queries.Add TbEstP.QrySave
         Queries.Add .QrySave
         
         
         bOk = mvarSys.xDb.Executa(Queries)
         
'         If bOk Then
'            If .TPTRANSA = "T" Then
'               nCtDe = mvarMe.CmbDe.ItemData(mvarMe.CmbDe.ListIndex)
'               nCtPara = mvarMe.CmbPara.ItemData(mvarMe.CmbPara.ListIndex)
'
'               mvaridmov = .IDTRANSF
'               mvarIDEST = IIf(.IDEST = nCtDe, nCtPara, nCtDe)
'               Call PopulaClasse(TbTransf)
'               mvarIDPROD = .IDPROD
'               mvarIDEST = .IDEST
'
'               TbTransf.IDTRANSF = .IDPROD
'               If TbTransf.isDirt Then
'                  bOk = TbTransf.Salvar
'               End If
'               If bOk Then
'                  .IDTRANSF = TbTransf.IDPROD
'                  If .isDirt And bOk Then
'                     bOk = .Salvar
'                  End If
'               End If
'            End If
'         End If
      End If
         
      If .IDFOR = 0 Then Call SalvarFavorecido(TbSMOVEST)

      If bOk And bIsDirt Then
         Call PopulaGrdLan
         If RsLan.EOF <> RsLan.BOF Then
            RsLan.MoveFirst
            RsLan.Find "IDMOV=" & .IDMOV
         End If
         nIndex = RsLan.AbsolutePosition - 1
         
         If nIndex >= 0 Then
            With mvarMe.GrdMov
               Set oRptRow = .Rows.FindRow(mvarMe.GrdMov.Records(nIndex))
               If Not oRptRow Is Nothing Then
                  oRptRow.Selected = True
                  .TopRowIndex = nTopRow
                  oRptRow.Selected = True
                  oRptRow.EnsureVisible
               End If
               .Redraw
               If .Visible And .Enabled Then
                  If bNovo Then
                     mvarMe_CmdNovoClick
                  Else
                     .SetFocus
                     Call HabilitaTab(False)
                     Call PopulaTab
                  End If
               End If
            End With
         End If
      End If
   End With

Saida:
   Call ExibirResultado(mvarSys, bOk, IIf(bOk, 1, 2))
'   Salvar = bOk
End Sub
Private Function ValidaCampos() As Boolean
   On Error GoTo TrataErro
   
   If Not IsDate(mvarMe.TxtDTMOV.Text) Then
      Call ExibirAviso("Data inválida.")
      mvarMe.TxtDTMOV.SetFocus
      GoTo Saida
   End If
   If xVal(mvarMe.TxtValor.Text) = 0 Then
      If vbNo = ExibirPergunta("Valor do lançamento iqual à zero." & vbNewLine & vbNewLine & "Deseja continuar com valor zerado?", , False) Then
         DoEvents
         mvarMe.TxtValor.SetFocus
         GoTo Saida
      End If
   End If
   If Trim(mvarMe.CmbFavorecido.Text) = "" Then
      Call ExibirAviso("Favorecido inválido.")
      mvarMe.CmbFavorecido.SetFocus
      GoTo Saida
   ElseIf mvarMe.CmbFavorecido.ListIndex < 0 Then
      'If vbYes = ExibirPergunta("Deseja Salvar o Favorecido '" & mvarMe.CmbFavorecido.Text & "'?", "Favorecido") Then
         Call SalvarFavorecido
         Call LocalizarCombo(mvarMe.CmbFavorecido, mvarMe.CmbFavorecido.Text)
      'End If
   End If
   If mvarMe.CmbFavorecido.ListIndex < 0 Then
      Call ExibirAviso("Favorecido inválido.")
      mvarMe.CmbFavorecido.SetFocus
      GoTo Saida
   End If
'   If mvarMe.TabConta.SelectedItem = 1 Then 'Tranferência
'      If Trim(mvarMe.CmbDe.Text) = "" Then
'         Call ExibirAviso("Conta origem inválida.")
'         mvarMe.CmbDe.SetFocus
'         GoTo Saida
'      ElseIf mvarMe.CmbDe.ListIndex < 0 Then
'         Call LocalizarCombo(mvarMe.CmbDe, mvarMe.CmbDe.Text)
'      End If
'      If mvarMe.CmbDe.ListIndex < 0 Then
'         Call ExibirAviso("Conta origem inválida.")
'         mvarMe.CmbDe.SetFocus
'         GoTo Saida
'      End If
'
'      If Trim(mvarMe.CmbPara.Text) = "" Then
'         Call ExibirAviso("Conta destino inválida.")
'         mvarMe.CmbPara.SetFocus
'         GoTo Saida
'      ElseIf mvarMe.CmbPara.ListIndex < 0 Then
'         Call LocalizarCombo(mvarMe.CmbPara, mvarMe.CmbPara.Text)
'      End If
'      If mvarMe.CmbPara.ListIndex < 0 Then
'         Call ExibirAviso("Conta destino inválida.")
'         mvarMe.CmbPara.SetFocus
'         GoTo Saida
'      End If
'
'      If mvarMe.CmbDe.ItemData(mvarMe.CmbDe.ListIndex) = mvarMe.CmbPara.ItemData(mvarMe.CmbPara.ListIndex) Then
'         Call ExibirAviso("Conta destino não pode ser iguaa a conta origem.")
'         mvarMe.CmbPara.SetFocus
'         GoTo Saida
'      End If
'   End If
                     
   ValidaCampos = True
TrataErro:
Saida:
End Function
Private Sub SalvarFavorecido(Optional TbMOV As Object)
   Dim TbContato As Object
   Dim bRefresh  As Boolean
   Dim nIDCont   As Long
   
   If Trim(mvarMe.CmbFavorecido.Text) <> "" Then
            
      Set TbContato = CriarObjeto("BANCO_3R.TB_OCONTATO")
      bRefresh = False
      
      With TbContato
         Set .xDb = mvarSys.xDb
         Call .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="NOME=" & SqlStr(mvarMe.CmbFavorecido.Text))
         .IDLOJA = mvarIDLOJA
         .NOME = mvarMe.CmbFavorecido.Text
         .IDCOLIGADA = mvarSys.Propriedades("IDCOLIGADA")
         .ATIVO = 1
         If .DTCADASTRO = "" Then .DTCADASTRO = mvarSys.xDb.SysDate
         If .isDirt Then
            If .Salvar Then
               bRefresh = True
            End If
         ElseIf .IDCONTATO <> 0 Then
            bRefresh = True
         End If
         If bRefresh Then
            If LocalizarCombo(mvarMe.CmbFavorecido, .IDCONTATO, False, True) < 0 Then
               mvarMe.CmbFavorecido.AddItem .NOME
               mvarMe.CmbFavorecido.ItemData(mvarMe.CmbFavorecido.NewIndex) = .IDCONTATO
               mvarMe.CmbFavorecido.ListIndex = mvarMe.CmbFavorecido.NewIndex
            End If
         End If
      End With
   End If
   
   If Not TbMOV Is Nothing Then
      With TbMOV
         If .IDFOR = 0 And mvarMe.CmbFavorecido.ListIndex >= 0 Then
            nIDCont = mvarMe.CmbFavorecido.ItemData(mvarMe.CmbFavorecido.NewIndex)
            
            .IDFOR = nIDCont
            '.TPFAVORECIDO = IIf(TbContato.PJ = 1, "F", "C")
            If .Salvar Then
               mvarMe.CmbFavorecido.ListIndex = -1
               Call LocalizarCombo(mvarMe.CmbFavorecido, CStr(nIDCont), True, True)
            End If
         End If
      End With
   End If
End Sub
Private Sub ExcluirLAN()
   Dim TbSMOVEST  As Object
   Dim TbTran  As Object
   Dim bOk     As Boolean
   Dim Queries As Collection
   Dim Sql     As String
         
   If vbNo = ExibirPergunta("Deseja realmente excluir o lançamento.", , False) Then GoTo Saida
         
   Set Queries = New Collection
   If VerificaLAN(bDelete:=True) Then
      Set TbSMOVEST = CriarObjeto("BANCO_3R.TB_SMOVEST")
      Set TbSMOVEST.xDb = mvarSys.xDb
      
      bOk = True

      If TbSMOVEST.Pesquisar(Ch_IDMOV:=mvarIDMOV) Then
         Queries.Add TbSMOVEST.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDMOV:=mvarIDMOV)
         
         If TbSMOVEST.IDTRANSF <> 0 Then
            Set TbTran = CriarObjeto("BANCO_3R.TB_SMOVEST")
            Set TbTran.xDb = mvarSys.xDb
            
            bOk = True
            If TbTran.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPROD:=TbSMOVEST.IDTRANSF) Then
               Queries.Add TbTran.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEST:=TbTran.IDEST, Ch_IDPROD:=TbTran.IDTRANSF)
            End If
         End If
         
         bOk = mvarSys.xDb.Executa(Queries)
      End If
      Set TbSMOVEST = Nothing
      Set Queries = Nothing
      
      If bOk Then
         Call PopulaGrdLan
         Call PopulaTab
      End If
      Call ExibirResultado(mvarSys, True, IIf(bOk, 1, 2))
   End If
Saida:

End Sub
Private Function VerificaLAN(Optional bDelete As Boolean, Optional pIDMOV As Long) As Boolean
   Dim nIDVENDA   As Long
   Dim nValor     As Double
   Dim sNUMDOC    As String
   Dim bOk        As Boolean
   Dim Sql        As String
   
   Dim bPossuiVenda  As Boolean
   Dim bAlteraValor  As Boolean
   Dim bAlteraNUMDOC As Boolean
   Dim TbItemVenda   As Object
   
   
   If mvarMe.GrdMov.SelectedRows(0).Index >= mvarMe.GrdMov.Records.Count - 1 Then
      If Not bDelete Then VerificaLAN = True
      Exit Function
   End If
   With mvarMe.GrdMov.SelectedRows(0).Record
      If pIDMOV = 0 And bDelete Then
         pIDMOV = .Item(iCampo("IDMOV")).Value
      End If
      If pIDMOV <> 0 Then
         nIDVENDA = xVal(.Item(iCampo("IDVENDA")).Value)
         nValor = xVal(.Item(iCampo("QTDITEM")).Value)
         sNUMDOC = .Item(iCampo("NUMDOC")).Value
      End If
      bPossuiVenda = (nIDVENDA <> 0)
      bAlteraValor = (Abs(nValor) <> xVal(mvarMe.TxtValor.Text))
      bAlteraNUMDOC = (sNUMDOC <> mvarMe.TxtNUMDOC.Text)
   End With
   
   bOk = True
   If bDelete Then
      If bPossuiVenda Then
         Sql = ""
         Sql = Sql & " And IDVENDA=" & nIDVENDA
         Sql = Sql & " And IDPROD=" & mvarIDPROD
         
         Set TbItemVenda = CriarObjeto("BANCO_3R.TB_CITENSVENDA")
         Set TbItemVenda.xDb = mvarSys.xDb
         
         If TbItemVenda.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=Sql) Then
            bOk = False
            Call ExibirAviso("Lançamento associado a venda '" & nIDVENDA & "' e não pode ter seu valor alterado ou excluído.")
         End If
      End If
   Else
      If bPossuiVenda And (bAlteraValor Or bAlteraNUMDOC) Then
         bOk = False
         Call ExibirAviso("Lançamento associado a venda '" & nIDVENDA & "' e não pode ter seu valor alterado ou excluído.")
         mvarMe.TxtNUMDOC.Text = sNUMDOC
         mvarMe.TxtValor.Text = ValBr(nValor, 0)
      End If
   End If
   VerificaLAN = bOk
End Function
Private Sub PopulaClasse(ByRef TbSMOVEST As Object, Optional ByRef Queries As Collection)
   Dim sDate   As String
   Dim Sql     As String
   Dim TbTran  As Object
   
   Set TbSMOVEST = CriarObjeto("BANCO_3R.TB_SMOVEST")
   With TbSMOVEST
      .xDb = mvarSys.xDb
      
      bNovo = Not .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDMOV:=mvarIDMOV)
      If bNovo Then
         .CODSIT = ""
         .DTCADASTRO = mvarSys.xDb.SysDate(3)
      End If
      .IDLOJA = mvarIDLOJA
      .IDEST = mvarIDEST
      .IDPROD = mvarIDPROD
      
      If Format(.DTMOV, "dd/mm/yyyy") <> mvarMe.TxtDTMOV.Text Then
         .DTMOV = mvarMe.TxtDTMOV.Text & " 23:59"
      End If
      .IDFOR = 0
      If mvarMe.CmbFavorecido.ListIndex >= 0 Then
         .IDFOR = mvarMe.CmbFavorecido.ItemData(mvarMe.CmbFavorecido.ListIndex)
      End If
      If .IDFOR = 0 Then .IDFOR = .IDFOR
      .FAVORECIDO = mvarMe.CmbFavorecido.Text
      .OBS = mvarMe.TxtObs.Text
      .NUMDOC = mvarMe.TxtNUMDOC.Text
      
      If xVal(mvarMe.TxtValor.Text) < 0 Then
         If .TPTRANSA = "C" Then mvarMe.TabConta.SelectedItem = 2
         If .TPTRANSA = "T" Then mvarMe.TabConta.SelectedItem = 1
         If .TPTRANSA = "D" Then mvarMe.TabConta.SelectedItem = 0
         If .TPTRANSA = "Q" Then mvarMe.TabConta.SelectedItem = 0
         
         mvarMe.TxtValor.Text = ValBr(-1 * xVal(mvarMe.TxtValor.Text), 0)
      End If
      If mvarMe.TabConta.SelectedItem = 0 Then '* Entrada
         .QTDITEM = xVal(mvarMe.TxtValor.Text)
      ElseIf mvarMe.TabConta.SelectedItem = 2 Then '* Saída
         .QTDITEM = xVal(mvarMe.TxtValor.Text) * -1
      End If
      If mvarMe.TabConta.SelectedItem = 0 Then .TPTRANSA = "C"
      If mvarMe.TabConta.SelectedItem = 1 Then .TPTRANSA = "T"
      If mvarMe.TabConta.SelectedItem = 2 Then .TPTRANSA = "D"
      If mvarMe.TabConta.SelectedItem = 3 Then .TPTRANSA = "Q"
      
      If .TPTRANSA = "T" Then
'         .TPLAN = IIf(mvarIDEST = mvarMe.CmbDe.ItemData(mvarMe.CmbDe.ListIndex), "D", "C")
      Else
         .TPLAN = IIf(mvarMe.TabConta.SelectedItem = 0, "C", "D")
      End If
      If .TPDOC <> "CV" Then
         If .NUMDOC <> "" And .TPTRANSA <> "T" Then
            If .TPLAN = "C" Then
               .TPDOC = "CV"
            Else
               .TPDOC = "NF"
            End If
         End If
      End If
      .OBS = mvarMe.TxtObs.Text
      .UNIDCONTROLE = mvarUNIDCONTROLE
      .FLGDELETE = 0
   End With
End Sub
Private Sub mvarMe_CommandBarsExecute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   If Control.Category = "Popup_IDPROD" Then
      mvarSys.Propriedades("IDPROD") = Control.Id
      mvarIDPROD = Control.Id
      Call WriteIniFile(mvarSys.LocalReg, "Config", "IDPROD", CStr(mvarIDPROD))
      Call PopulaTela
      mvarMe.SccTit.Caption = "  Produto: " & Control.Caption
   End If
End Sub

Private Sub mvarMe_GrdMovBeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
   If Not iCampo Is Nothing Then
      'If Item.Index = iCampo("Entrada") Or Item.Index = iCampo("Saída") Or Item.Index = iCampo("Saldo") Then
      '   Item.Value = IIf(xVal(Item.Value & "") = 0, "", ValBr(Item.Value & ""),0)
      'End If
   End If
   If ((Row.Index Mod 2 = 0) And (Row.GroupRow = False)) Then
       'This will dynamically change the BackColor of the ReportRecordItem, this value
       'is not saved in the Item.Backcolor property.  Once mnuAutomaticFormating.Checked is false,
       'the original backcolor is restored because this code is not executed
       Metrics.BackColor = RGB(245, 245, 245)
   Else
      'Metrics.BackColor = RGB(65, 165, 241)
   End If
End Sub
Private Sub mvarMe_GrdMovRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   On Error Resume Next
   Call HabilitaTab(True)
   Select Case Item.Index
      Case iCampo("Núm"): mvarMe.TxtNUMDOC.SetFocus
      Case iCampo("Data"): mvarMe.TxtDTMOV.SetFocus
      Case iCampo("Fornecedor/Cliente"): mvarMe.CmbFavorecido.SetFocus
      Case iCampo("Saída"): mvarMe.TxtValor.SetFocus
      Case iCampo("Entrada"): mvarMe.TxtValor.SetFocus
      Case iCampo("Saldo"): mvarMe.TxtValor.SetFocus
      Case Else:
         mvarMe.TxtNUMDOC.SetFocus
         If Trim(mvarMe.TxtNUMDOC.Text) = "" Then
            SendKeys "{Tab}"
         End If
   End Select
End Sub
Private Sub mvarMe_GrdMovSelectionChanged()
   Call LimpaTab
   Call HabilitaTab(False)
   Call PopulaTab
   If Not mvarMe.CmbFavorecido.Visible Then
      Call mvarMe_TabContaSelectedChanged(mvarMe.TabConta.Item(mvarMe.TabConta.SelectedItem))
   End If
   
   DoEvents
   If mvarMe.GrdMov.Visible And mvarMe.GrdMov.Enabled Then
      mvarMe.GrdMov.SetFocus
   End If
End Sub
Private Sub LimpaTab()
   With mvarMe
'      Call LocalizarCombo(.CmbDe, "")
'      Call LocalizarCombo(.CmbPara, "")
      Call LocalizarCombo(.CmbFavorecido, "")
      .TxtObs.Text = ""
      .TxtNUMDOC.Text = ""
      .TxtDTMOV.Text = ""
      .TxtValor.Text = ""
   End With
End Sub
Private Sub HabilitaTab(pEnabled As Boolean)
'   If pEnabled Then
'      Call mvarMe_TabContaSelectedChanged(mvarMe.TabConta.Item(mvarMe.TabConta.SelectedItem))
'   End If
   
   With mvarMe
      If Not pEnabled Then
'         .CmbDe.Text = ""
'         .CmbPara.Text = ""
         .CmbFavorecido.Text = ""
      End If
      
'      .CmbDe.Enabled = pEnabled
'      .CmbPara.Enabled = pEnabled
      .CmbFavorecido.Enabled = pEnabled
      .TxtObs.Enabled = pEnabled
      .TxtNUMDOC.Enabled = pEnabled
      .TxtDTMOV.Enabled = pEnabled
      .TxtValor.Enabled = pEnabled
      
'      .CmbDe.Visible = True
'      .CmbPara.Visible = True
      .CmbFavorecido.Visible = True
      .TxtObs.Visible = True
      .TxtNUMDOC.Visible = True
      .TxtDTMOV.Visible = True
      .TxtValor.Visible = True
      
      
      mvarMe.CmdSalvar.Enabled = pEnabled
      mvarMe.CmdCancelar.Enabled = pEnabled
      mvarMe.CmdNovo.Enabled = Not pEnabled
   End With
End Sub
Private Sub PopulaTab()
   Dim nPag As Currency
   Dim nDep As Currency
   Dim oReg As ReportRecord
   Dim Sql As String
   
   With mvarMe
      If .GrdMov.SelectedRows.Count <> 1 Then Exit Sub

      Set oReg = .GrdMov.SelectedRows(0).Record

      mvarIDMOV = RegLan("IDMOV").Value
            
      bNovo = (mvarIDMOV = 0)
      Select Case RegLan("TPTRANSA").Value
         Case "C": .TabConta.SelectedItem = 0
         Case "T": .TabConta.SelectedItem = 1
         Case "D": .TabConta.SelectedItem = 2
         Case Else: If mvarIDMOV <> 0 Then .TabConta.SelectedItem = 0
      End Select
      
'      If RegLan("TPTRANSA").Value = "T" Then
'
'         Dim TbTransf   As Object
'         Dim nIDLOJA    As Long
'         Dim nIDEST_T As Long
'         Dim nIDTRANSF  As Long
'
'         nIDLOJA = RegLan("IDLOJA").Value
'         nIDTRANSF = RegLan("IDTRANSF").Value
'
'         nIDEST_T = xVal(GetTag(RegLan("TPLAN"), "IDEST_T", 0))
'         If nIDEST_T = 0 Then
'            Set TbTransf = CriarObjeto("BANCO_3R.TB_SMOVEST")
'            With TbTransf
'               Set .xDb = mvarSys.xDb
'               If .Pesquisar(Ch_IDLOJA:=nIDLOJA, Ch_IDLAN:=nIDTRANSF) Then
'                  nIDEST_T = .IDCONTA
'                  Call SetTag(RegLan("TPLAN"), "IDEST_T", nIDEST_T)
'               End If
'            End With
'            Set TbTransf = Nothing
'         End If
'         If RegLan("TPLAN").Value = "D" Then
'            Call LocalizarCombo(.CmbDe, RegLan("IDEST").Value, True, True)
'            Call LocalizarCombo(.CmbPara, CStr(nIDEST_T), True, True)
'         Else
'            Call LocalizarCombo(.CmbDe, CStr(nIDEST_T), True, True)
'            Call LocalizarCombo(.CmbPara, RegLan("IDEST").Value, True, True)
'         End If
'      End If
      .TxtNUMDOC.Text = RegLan("Núm").Value
      .TxtDTMOV.Text = RegLan("Data").Value
      If IsDate(.TxtDTMOV.Text) Then
         .CmbDTMOV.Value = CDate(.TxtDTMOV.Text)
      End If
      
      nPag = xVal(RegLan("Saída").Value)
      nDep = xVal(RegLan("Entrada").Value)
      .TxtValor.Text = ValBr(nPag + nDep, 0)
      
      .CmbFavorecido.ListIndex = -1
      
      .CmbFavorecido.ListIndex = LocalizarCombo(.CmbFavorecido, RegLan("IDFOR").Value, False, True)
      If InStr(RegLan("Fornecedor/Cliente").Value, ":") <> 0 Then
         .CmbFavorecido.Text = Mid(RegLan("Fornecedor/Cliente").Value, 1, InStr(RegLan("Fornecedor/Cliente").Value, ":") - 1)
      Else
         .CmbFavorecido.Text = RegLan("Fornecedor/Cliente").Value
      End If
      If .CmbFavorecido.Text = "" And xVal(.TxtValor.Text) > 0 Then
         Sql = "SET IDENTITY_INSERT OCONTATO on;" & vbNewLine
         Sql = Sql & " Insert OCONTATO (IDLOJA,IDCOLIGADA, IDCONTATO, NOME) Values (" & mvarIDLOJA & "," & mvarSys.Propriedades("IDCOLIGADA") & ",1,'Cliente');" & vbNewLine
         Sql = Sql & " SET IDENTITY_INSERT OCONTATO off;" & vbNewLine
         Call mvarSys.xDb.Executa(Sql)
      End If
      .TxtObs.Text = RegLan("OBS").Value
      
      If .TxtDTMOV.Text = "" Then .TxtValor.Text = ""
   End With
End Sub

Private Sub mvarMe_ImgMenuMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Call ExibirMenuPROD(Button, Shift)
End Sub

Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
'   DoEvents
   If Shift = 4 Then 'vbKeyControl
      If KeyCode = vbKeyD Then mvarMe.TabConta.SelectedItem = 0
      If KeyCode = vbKeyT Then mvarMe.TabConta.SelectedItem = 1
      If KeyCode = vbKeyR Then mvarMe.TabConta.SelectedItem = 2
      If KeyCode = vbKeyC Then mvarMe.TabConta.SelectedItem = 3
   End If
   If KeyCode = vbKeyF5 Then
      Call PopulaGrdLan
   ElseIf KeyCode = vbKeyDelete Then
      If mvarMe.ActiveControl Is mvarMe.GrdMov Then
         Call ExcluirLAN
      End If
   
   ElseIf KeyCode = vbKeyEscape Then
      If mvarMe.txtFiltrar.Text <> StrPesquisa Then
         mvarMe.txtFiltrar.Text = StrPesquisa
         Call PopulaGrdLan
      End If
      Call PopulaTab
      
   ElseIf KeyCode = vbKeyInsert Then
      Call mvarMe_CmdNovoClick
      
   ElseIf (Asc(UCase(Chr(KeyCode))) >= Asc("A") And Asc(UCase(Chr(KeyCode))) >= Asc("Z")) Or (Asc(UCase(Chr(KeyCode))) >= Asc("0") And Asc(UCase(Chr(KeyCode))) >= Asc("9")) Then
      'If Not mvarMe.ActiveControl Is mvarMe.txtFiltrar Then
      '   mvarMe.txtFiltrar.Text = UCase(Chr(KeyCode))
      '   mvarMe.txtFiltrar.SetFocus
      '   mvarMe.txtFiltrar.SelStart = Len(mvarMe.txtFiltrar.Text)
      'End If
   End If
End Sub
Private Sub mvarMe_Load()
   mvarMe.CommandBars.DeleteAll
   Call MontaTela
   Call PopulaTela
   Call mvarMe_Rezise
   DoEvents
End Sub
Public Sub PopulaTela()
   Call LimpaGrd
   Call PopulaGrdLan
   'mvarMe_Rezise
   'Call mvarMe_TabContaSelectedChanged(mvarMe.TabConta.Item(mvarMe.TabConta.SelectedItem))
End Sub
Private Sub MontaTela()
   'ReportControlGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   
   With mvarMe
      .TxtObs.Visible = False
      
      .GrpPg.ZOrder 1
      
      .txtFiltrar = StrPesquisa
      .SccTit.VisualTheme = xtpShortcutThemeOffice2003
      .SccTit.GradientHorizontal = False
      
      With .GrdMov
         .BorderStyle = xtpBorderFrame
         .PaintManager.ColumnStyle = xtpColumnOffice2007
         .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
         .PaintManager.RefreshMetrics
         
         .SetImageList mvarMe.ImageList1
         .SetCustomDraw xtpCustomBeforeDrawRow
         
         .MultipleSelection = False
         .AllowColumnRemove = False
         .AllowColumnReorder = False
         .AllowColumnResize = False
         .AllowColumnSort = False
         .AllowEdit = False
                     
         .FooterRowsAllowAccess = True
         .FooterRowsAllowEdit = True
         .FooterRowsEnableSelection = True
      End With
      
      .TabConta.Item(1).Visible = False
      
      '**********************************
      '* Estoques
      '**********************************
      Dim MyTbEst As Object
      
'      .CmbDe.AutoComplete = True
'      .CmbPara.AutoComplete = True
'      .CmbDe.Clear
'      .CmbPara.Clear
      
'      Set MyTbEst = CriarObjeto("BANCO_3R.TB_SESTOQUE")
'      With MyTbEst
'         .xdb = mvarSys.xdb
'         If .Pesquisar() Then
'            While Not .RS.EOF
'               .Popula
'               mvarMe.CmbDe.AddItem .DSCEST
'               mvarMe.CmbDe.ItemData(mvarMe.CmbDe.NewIndex) = .IDEST
'               mvarMe.CmbPara.AddItem .DSCEST
'               mvarMe.CmbPara.ItemData(mvarMe.CmbPara.NewIndex) = .IDEST
'               .RS.MoveNext
'            Wend
'         End If
'      End With
      
      '**********************************
      '* Favorecido
      '**********************************
      Dim Sql As String
      Dim MyRs As Object
      
      .CmbFavorecido.AutoComplete = True
      .CmbFavorecido.Sorted = True
      .CmbFavorecido.Clear
      
      Sql = "Select Distinct C.IDCONTATO, C.NOME"
      Sql = Sql & " From OCONTATO C, SMOVEST M"
      Sql = Sql & " Where C.IDCONTATO=M.IDFOR"
      Sql = Sql & " And C.ATIVO=1"
      Sql = Sql & " And C.IDLOJA=" & SqlNum(mvarIDLOJA)
      
      Sql = "Select Distinct C.IDCONTATO, C.NOME"
      Sql = Sql & " From OCONTATO C"
      Sql = Sql & " Where C.ATIVO=1"
      Sql = Sql & " And C.IDLOJA=" & SqlNum(mvarIDLOJA)
      Sql = Sql & " Order By C.NOME"
      
      If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
         While Not MyRs.EOF
            .CmbFavorecido.AddItem MyRs("NOME")
            .CmbFavorecido.ItemData(.CmbFavorecido.NewIndex) = MyRs("IDCONTATO")
            MyRs.MoveNext
         Wend
      End If
      
      Call MontaProduto
      Call SelecionaProduto
   End With
End Sub
Private Sub SelecionaProduto()
   Dim Sql As String
   
   Sql = "Select IDPROD, NMPROD, CODPROD, UNIDCONTROLE "
   Sql = Sql & " From SPRODUTO "
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And IDPROD=" & mvarIDPROD
   Sql = Sql & " And ESERVICO=0 "
   Sql = Sql & " And SITPROD='A'"
   
   If mvarSys.xDb.AbreTabela(Sql) Then
      mvarMe.SccTit.Caption = "  Produto: " & mvarSys.xDb.RsAux("NMPROD")
   Else
      Sql = "Select IDPROD, NMPROD, CODPROD, UNIDCONTROLE "
      Sql = Sql & " From SPRODUTO "
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And ESERVICO=0 "
      Sql = Sql & " And SITPROD='A'"
      Sql = Sql & " Order By EVENDA Desc"
      If Not mvarSys.xDb.AbreTabela(Sql) Then
         mvarMe.SccTit.Caption = "  Produto: " & mvarSys.xDb.RsAux("NMPROD")
         mvarIDPROD = mvarSys.xDb.RsAux("IDPROD")
         mvarUNIDCONTROLE = mvarSys.xDb.RsAux("UNIDCONTROLE")
         mvarMe.SccTit.Caption = "  Produto: " & mvarSys.xDb.RsAux("NMPROD") '& " (Cod." & mvarSys.xDb.RsAux("CODPROD") & ")"
         mvarMe.SccTit.ForeColor = vbWhite
         mvarMe.SccTit.Font.Size = 13
      Else
         mvarIDPROD = 0
         mvarUNIDCONTROLE = ""
         mvarMe.SccTit.Caption = "  Por favor selecione um produto"
         mvarMe.SccTit.ForeColor = &H8000&
         mvarMe.SccTit.Font.Size = 13
      End If
   End If
End Sub
Private Sub MontaProduto()
   Dim Sql As String
   Dim TbProd As Object

   With mvarMe
      .CmbProduto.Clear

      .CmbProduto.AutoComplete = True
      .CmbProduto.BackColor = mvarMe.SccTit.GradientColorDark
      .CmbProduto.UseVisualStyle = True
      .CmbProduto.Visible = False
      
      If .CmbProduto.Visible Then
         Sql = ""
         Sql = Sql & "SITPROD='A'"
         Sql = Sql & " And EVENDA=1"
         Sql = Sql & " And ESERVICO=0"
               
         Set TbProd = CriarObjeto("BANCO_3R.TB_SPRODUTO")
         With TbProd
            .xDb = mvarSys.xDb
            If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=Sql) Then
               While Not .RS.EOF
                  .Popula
                  mvarMe.CmbProduto.AddItem .NMPROD
                  mvarMe.CmbProduto.ItemData(mvarMe.CmbProduto.NewIndex) = .IDPROD
                  .RS.MoveNext
               Wend
               If mvarMe.CmbProduto.ListCount = 1 Then
                  mvarMe.CmbProduto.ListIndex = 0
               End If
            End If
         End With
      End If
   End With
End Sub
Private Sub mvarMe_Rezise()
   On Error Resume Next
   mvarMe.GrpPg.ZOrder 1

   mvarMe.SccTit.Move 0, 50, mvarMe.ScaleWidth - 50, 420
   With mvarMe.txtFiltrar
      .Left = mvarMe.Width - 2800
      .Top = 90
      .Visible = False
   End With
   With mvarMe.imgLupa
      .Left = mvarMe.Width - 680
      .Top = 100
      .Height = mvarMe.txtFiltrar.Height - 25
      .Visible = False
   End With
   With mvarMe.CmbProduto
      .Top = 100
      .Left = mvarMe.Width - .Width - 120
   End With
   
   With mvarMe.TabConta
      .Left = 0
      .Top = mvarMe.ScaleHeight - .Height - 60
      .Width = mvarMe.ScaleWidth - 50
   End With
   With mvarMe.GrpPg
      .Top = mvarMe.TabConta.Top + mvarMe.TabPgDeposito.Top
      .Left = mvarMe.TabConta.Left + 30
      .Width = mvarMe.TabConta.Width - 60    ' mvarMe.TabPgDeposito.Width
      .Height = mvarMe.TabConta.Height - 390 ' mvarMe.TabPgDeposito.Height
   End With
   mvarMe.GrpValor.Top = 0
   mvarMe.GrpValor.Left = mvarMe.GrpPg.Width - mvarMe.GrpValor.Width
   mvarMe.GrpDe.Top = 120
   mvarMe.GrpDe.Width = mvarMe.GrpValor.Left - mvarMe.GrpDe.Left
   mvarMe.GrpDe.Height = mvarMe.GrpPg.Height - mvarMe.GrpDe.Top
      
   With mvarMe.GrdMov
      .Left = 0
      .Top = mvarMe.SccTit.Top + mvarMe.SccTit.Height + 30
      .Width = mvarMe.ScaleWidth - 50
      .Height = mvarMe.TabConta.Top - .Top - 30
   End With
   
   mvarMe.CmbDTMOV.Top = mvarMe.TxtDTMOV.Top
   mvarMe.CmbDTMOV.Height = mvarMe.TxtDTMOV.Height
'   mvarMe.CmbDe.Width = mvarMe.GrpValor.Left - mvarMe.CmbDe.Left - 320
'   mvarMe.CmbPara.Width = mvarMe.GrpValor.Left - mvarMe.CmbPara.Left - 320
   mvarMe.CmbFavorecido.Width = mvarMe.GrpValor.Left - mvarMe.CmbFavorecido.Left - 320
   mvarMe.TxtObs.Width = mvarMe.GrpValor.Left - mvarMe.TxtObs.Left - 320
            
   mvarMe.CmdNovo.Top = 90
   mvarMe.CmdNovo.Height = 300
   mvarMe.CmdSalvar.Top = mvarMe.CmdNovo.Top
   mvarMe.CmdCancelar.Top = mvarMe.CmdNovo.Top
   mvarMe.CmdSalvar.Height = mvarMe.CmdNovo.Height
   mvarMe.CmdCancelar.Height = mvarMe.CmdNovo.Height
   
   mvarMe.LblTitulo.Caption = mvarMe.TabConta.Item(mvarMe.TabConta.SelectedItem).Caption
   mvarMe.LblTitulo.Left = mvarMe.TabConta.Width - mvarMe.LblTitulo.Width - 240
            
   Call DefineGrdAlinha
      
   mvarMe.GrpPg.ZOrder 0
End Sub
Private Sub mvarMe_SccTitMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Call ExibirMenuPROD(Button, Shift)
End Sub
Private Sub ExibirMenuPROD(Button As Integer, Shift As Integer)
   Dim Popup   As CommandBar
   Dim oCmd    As CommandBarControl
   Dim Sql As String
   Dim oRs As Object
   
   If Button = 2 Then
      Sql = "Select Distinct P.IDPROD, P.NMPROD" & vbNewLine
      Sql = Sql & " From SPRODUTO P" & vbNewLine
      'Sql = Sql & " Join SMOVEST M On M.IDLOJA=P.IDLOJA And M.IDPROD=P.IDPROD" & vbNewLine
      Sql = Sql & " Where P.IDLOJA=" & mvarIDLOJA & vbNewLine
      Sql = Sql & " And P.ESERVICO=0" & vbNewLine
      Sql = Sql & " And P.SITPROD='A'" & vbNewLine
      'Sql = Sql & " And P.EVENDA=1 "
      Sql = Sql & " Order By P.NMPROD"
      If mvarSys.xDb.AbreTabela(Sql, oRs) Then
         Set Popup = mvarMe.CommandBars.Add("Popup", xtpBarPopup)
         With Popup
            .Controls.DeleteAll
            While Not oRs.EOF
               Set oCmd = .Controls.Add(XTPControlType.xtpControlButton, oRs("IDPROD"), oRs("NMPROD"))
               oCmd.Category = "Popup_IDPROD"
               oCmd.Checked = (mvarIDPROD = oRs("IDPROD"))
               oRs.MoveNext
            Wend
           .ShowPopup
           .Controls.DeleteAll
         End With
      End If
   End If
End Sub

Private Sub mvarMe_TabContaSelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   Dim Lin1 As Integer
   Dim Lin2 As Integer
   Dim Lin3 As Integer
   Dim Lin4 As Integer
   Dim Lin5 As Integer
   
   DoEvents
   Lin1 = 0
   Lin2 = 360
   Lin3 = 720
   Lin4 = 1080
   Lin5 = 1440
         
   With mvarMe
      If .CmdNovo.Enabled And .CmdNovo.Visible Then
         Select Case RegLan("TPTRANSA").Value
            Case "C": If Item.Index <> 0 Then Call mvarMe_CmdNovoClick
            Case "T": If Item.Index <> 1 Then Call mvarMe_CmdNovoClick
            Case "D": If Item.Index <> 2 Then Call mvarMe_CmdNovoClick
            Case "Q": If Item.Index <> 3 Then Call mvarMe_CmdNovoClick
         End Select
      End If
   
      .LblTitulo.AutoSize = True
      .LblTitulo.Font.Size = 16
      .LblTitulo.Caption = Item.Caption
      .LblTitulo.Left = .TabConta.Width - .LblTitulo.Width - 240
         
'      .LblDe.Visible = False
'      .LblPara.Visible = False
      .LblFavorecido.Visible = False
      .LblObs.Visible = False
      
      .LblFavorecido.Caption = "Pagar a:"
      
'      .GrdMov.TabIndex = 1
'      .TabConta.TabIndex = 2
'      .CmdNovo.TabIndex = 3
'      .CmdSalvar.TabIndex = 4
'      .CmdCancelar.TabIndex = 5
'
'      .LblNum.TabIndex = 6
'      .CmdNum.TabIndex = 7
'      .TxtNUMDOC.TabIndex = 8
      
      Select Case Item.Index
         Case 0, 2, 3 '* Depóstio
            If Item.Index = 0 Then .LblFavorecido.Caption = "Receber de:"
            '.LblDe.Top = Lin1
            '.LblPara.Top = Lin1
            .LblFavorecido.Top = Lin2
            .LblObs.Top = Lin3

'            .LblFavorecido.TabIndex = 9
'            .CmbFavorecido.TabIndex = 10
'            .LblDTMOV.TabIndex = 11
'            .TxtDTMOV.TabIndex = 12
'            .LblObs.TabIndex = 13
'            .TxtObs.TabIndex = 14
'            .LblVALOR.TabIndex = 15
'            .TxtValor.TabIndex = 16
            
                  
         Case 1 '* Transferência
'            .LblDe.Top = Lin2
'            .LblPara.Top = Lin3
            .LblFavorecido.Top = Lin4
            '.LblCategoria.Top = Lin4
            .LblObs.Top = Lin5
            
'            .LblDe.TabIndex = 9
'            .CmbDe.TabIndex = 10
'            .LblDTMOV.TabIndex = 11
'            .TxtDTMOV.TabIndex = 12
'            .LblPara.TabIndex = 13
'            .CmbPara.TabIndex = 14
'            .LblVALOR.TabIndex = 15
'            .TxtValor.TabIndex = 16
'            .LblFavorecido.TabIndex = 17
'            .CmbFavorecido.TabIndex = 18
'            .LblObs.TabIndex = 19
'            .TxtObs.TabIndex = 20
            
'            If Not .CmdNovo.Enabled And .CmbDe.ListIndex < 0 Then
'               Call SelContaTrans
'            End If
            
         Case 2 '* Retirada
         Case 3 '* Cheque
      End Select

'      .LblDe.Visible = InArray(Item.Index, Array(1))
'      .LblPara.Visible = InArray(Item.Index, Array(1))
      .LblFavorecido.Visible = InArray(Item.Index, Array(0, 1, 2, 3))
      .LblObs.Visible = InArray(Item.Index, Array(0, 1, 2, 3))
   
      .LblNum.Visible = InArray(Item.Index, Array(0, 1, 2, 3))
      .CmdNum.Visible = InArray(Item.Index, Array(0, 1, 2, 3))
      .LblDTMOV.Visible = InArray(Item.Index, Array(0, 1, 2, 3))
      .LblVALOR.Visible = InArray(Item.Index, Array(0, 1, 2, 3))
      '===================================================
'      .CmbDe.Visible = .LblDe.Visible
'      .CmbPara.Visible = .LblPara.Visible
      .CmbFavorecido.Visible = .LblFavorecido.Visible
      .TxtObs.Visible = .LblObs.Visible
      
      .TxtNUMDOC.Visible = (.LblNum.Visible Or .CmdNum.Visible)
      .CmbDTMOV.Visible = .LblDTMOV.Visible
      .TxtDTMOV.Visible = .LblDTMOV.Visible
      .TxtValor.Visible = .LblVALOR.Visible
      '===================================================
      
'      .LblDe.Top = Lin1
'      .LblPara.Top = Lin2
'      .LblFavorecido.Top = Lin2
'      .LblCategoria.Top = Lin3
      
'      .CmbDe.Top = .LblDe.Top
'      .CmbPara.Top = .LblPara.Top
      .CmbFavorecido.Top = .LblFavorecido.Top
      .TxtObs.Top = .LblObs.Top
      
      '.Refresh
      '.GrpPg.ZOrder 1
      .GrpPg.ZOrder 0
   End With
End Sub
Private Sub mvarMe_TxtDTMOVLostFocus()
   mvarMe.TxtDTMOV.Text = FormatarData(mvarMe.TxtDTMOV.Text)
   'Call PopulaUltTransf
End Sub
Private Sub mvarMe_txtFiltrarGotFocus()
   mvarMe.txtFiltrar.ForeColor = vbBlack
   
   If mvarMe.txtFiltrar.Text = StrPesquisa Then
      mvarMe.txtFiltrar.Text = ""
   End If
End Sub

Private Sub mvarMe_txtFiltrarKeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      mvarMe.GrdMov.FilterText = mvarMe.txtFiltrar.Text
      mvarMe.GrdMov.Populate
   End If
End Sub

Private Sub mvarMe_txtFiltrarLostFocus()
   
   If mvarMe.txtFiltrar.Text = "" Then mvarMe.txtFiltrar.Text = StrPesquisa
   mvarMe.txtFiltrar.ForeColor = 12632256
End Sub
Private Sub mvarMe_TxtNUMDOCLostFocus()
   If xVal(mvarMe.TxtNUMDOC.Text) > 0 Then
      mvarMe.TxtNUMDOC.Text = StrZero(mvarMe.TxtNUMDOC.Text, mvarSys.Propriedades("LENIDVENDA"))
   End If
End Sub
Private Sub PopulaGrdLan(Optional pWhere As String, Optional pIDPROD As Integer = 0)
   Dim Sql As String
   Dim i As Long
   Dim nTopRow As Long
   Dim nRowSel As Long
   Dim nTabSel As Integer
   Dim oRecord As ReportRecord
   Dim oRptRow As ReportRow
   Dim nDias   As Integer
   Dim sDate   As String
   Dim bOk     As Boolean
   
   Sql = QueryGrid
   If Not mvarPane Is Nothing Then
      nDias = mvarPane.CmbPeriodo.ItemData(mvarPane.CmbPeriodo.ListIndex)
      If mvarPane.CmbPeriodo.ItemData(mvarPane.CmbPeriodo.ListIndex) > 0 Then
         sDate = mvarSys.xDb.SysDate()
         sDate = "01/" & Format(CDate(sDate) - nDias, "mm/yyyy")
         Sql = Sql & " And M.DTMOV>=" & SqlDate(sDate)
      End If
   End If
   Sql = Sql & IIf(pWhere = "", "", "AND " + pWhere)
   Sql = Sql & " Order By NLAN"
   
   Screen.MousePointer = vbHourglass
   nTabSel = mvarMe.TabConta.SelectedItem
   With mvarMe.GrdMov
      nTopRow = .TopRowIndex
      If .SelectedRows.Count <= 0 Then
         nRowSel = -1
      Else
         nRowSel = .SelectedRows(0).Index
         Set oRecord = .SelectedRows(0).Record
      End If
      
      .Visible = False
      bOk = False
      If mvarSys.xDb.AbreTabela(Sql, RsLan) Then
         bOk = (RsLan.RecordCount > 0)
      Else
         Sql = QueryGrid()
         If mvarSys.xDb.AbreTabela(Sql, RsLan) Then
            bOk = (RsLan.RecordCount > 0)
         End If
      End If
      If bOk Then
         Call L_FillRCFromRS(RsLan, mvarMe.GrdMov)
         Call EscondeCampos
      
         If nRowSel < 0 Then
            nRowSel = RsLan.RecordCount
         End If
      Else
         Sql = QueryGrid(True)
         Call mvarSys.xDb.AbreTabela(Sql, RsLan)
         Call FillRCFromRS(RsLan, mvarMe.GrdMov)
         Call EscondeCampos
         If nRowSel < 0 Then
            nRowSel = RsLan.RecordCount
         End If
         .InsertRecord .Rows.Count
         .Populate
         If .Rows.Count > 0 Then
            .Rows(.Rows.Count - 1).Selected = True
            Call PopulaTab
         End If
      
      End If
'      Call DefineGrdIcones
      Call DefineGrdFiltro
      Call CalculaSaldo
      Call DefineGrdAlinha
      
      Call InserirGrdLinha
      
      If pIDPROD <> 0 Then
         '.Rows.FindRow...
      End If
      '.Populate
      'DoEvents
      
      If nRowSel >= .Rows.Count Then nRowSel = .Rows.Count - 1
      If nTopRow >= .Rows.Count Then nTopRow = .Rows.Count - 1
      
      mvarMe.TabConta.SelectedItem = nTabSel
      '.Visible = True
      .TopRowIndex = nTopRow
      If Not oRecord Is Nothing Then
         Set oRptRow = .Rows.FindRow(oRecord)
      End If
      If oRptRow Is Nothing Then
         If nRowSel >= 0 Then
            .Rows(nRowSel).Selected = True
            .Rows(nRowSel).EnsureVisible
         End If
      Else
         oRptRow.Selected = True
         oRptRow.EnsureVisible
      End If
      
      .Visible = True
      .Redraw
      If .Visible And .Enabled Then
         .SetFocus
         If oRptRow Is Nothing Then
            If nRowSel >= 0 Then
               .Rows(nRowSel).Selected = True
            End If
         Else
            oRptRow.Selected = True
         End If
         
         
      End If
   End With
   Screen.MousePointer = vbDefault
End Sub
Private Sub InserirGrdLinha()
   mvarIDMOV = 0
   With mvarMe.GrdMov
      If .Rows.Count > 0 Then
         '.Rows(.Rows.Count - 1).EnsureVisible
         .Rows(.Rows.Count - 1).Selected = True
         Call PopulaTab
      
         If .Rows(.Rows.Count - 1).Record(iCampo("IDMOV")).Value > 0 Then
            .InsertRecord .Rows.Count
            .Populate
            If .Rows.Count > 0 Then
               '.Rows(.Rows.Count - 1).EnsureVisible
               .Rows(.Rows.Count - 1).Selected = True
               Call PopulaTab
            End If
         End If
      End If
   End With
End Sub
Private Sub CalculaSaldo()
   Dim vRec    As ReportRecord
   Dim nSaldo  As Double
   Dim nCred   As Double
   Dim nDeb    As Double
   Dim vCol    As ReportColumn
   
   With mvarMe.GrdMov
      nSaldo = 0
      
      .ShowFooter = True
      For Each vCol In .Columns
         vCol.FooterText = " "
         vCol.DrawFooterDivider = False
         vCol.DrawFooterRowsVGrid = False
      Next
      
      .Columns(iCampo("Saída")).FooterFont.Bold = True
      .Columns(iCampo("Saída")).FooterAlignment = xtpAlignmentCenter
      .Columns(iCampo("Saída")).FooterText = "Saldo Final :"
      
      .Columns(iCampo("Saldo")).FooterFont.Bold = True
      .Columns(iCampo("Saldo")).FooterAlignment = xtpAlignmentRight
      If .Records.Count > 0 Then
         .Columns(iCampo("Saldo")).FooterText = ValBr(.Records(.Records.Count - 1).Item(iCampo("Saldo")).Value, 0)
      End If
   End With
End Sub
Private Function QueryGrid(Optional bLimpaGrd As Boolean) As String
   Dim Sql As String
   Dim SqlSaldo As String
   Dim SqlMOV  As String
   Dim sTbRow As String
   Dim sTbRow2 As String
   
   SqlMOV = "Select M.IDLOJA , M.IDEST, M.IDPROD, M.DTMOV, M.IDMOV" & vbNewLine
   SqlMOV = SqlMOV & ", M.QTDITEM, M.UNIDCONTROLE, M.TPDOC, M.NUMDOC, M.ITEMDOC" & vbNewLine
   SqlMOV = SqlMOV & ", M.IDFOR, M.FLGDELETE, M.TPLAN, M.TPTRANSA, M.OBS" & vbNewLine
   SqlMOV = SqlMOV & ", M.IDVENDA, M.IDTRANSF" & vbNewLine
   If mvarSys.GetParam("FINNMFAVORECIDO") = 1 Then
      SqlMOV = SqlMOV & ", [FAVORECIDO]= IsNull(F.NOME, '')+Case IsNull(C.NOME, '')  WHEN '' THEN '' ELSE ': '+C.NOME End" & vbNewLine
      SqlMOV = SqlMOV & ", C.IDCLIENTE, C.NOME [NMCLIENTE]" & vbNewLine
   Else
      SqlMOV = SqlMOV & ", [FAVORECIDO]= IsNull(F.NOME, '')" & vbNewLine
   End If
   SqlMOV = SqlMOV & ", [CREDITO]   = Case WHEN M.QTDITEM>0 THEN M.QTDITEM    ELSE 0 End" & vbNewLine
   SqlMOV = SqlMOV & ", [DEBITO]     = Case WHEN M.QTDITEM<0 THEN -1*M.QTDITEM ELSE 0 End" & vbNewLine
   SqlMOV = SqlMOV & " From SMOVEST M " & vbNewLine
   SqlMOV = SqlMOV & " Left Join OCONTATO F On M.IDLOJA=F.IDLOJA And M.IDFOR=F.IDCONTATO" & vbNewLine
   If mvarSys.GetParam("FINNMFAVORECIDO") = 1 Then
      SqlMOV = SqlMOV & " Left Join CVENDA   V On V.IDLOJA=M.IDLOJA And V.IDVENDA=M.IDVENDA" & vbNewLine
      SqlMOV = SqlMOV & " Left Join OCLIENTE C On V.IDLOJA=C.IDLOJA And V.IDCLIENTE=C.IDCLIENTE" & vbNewLine
   End If
   SqlMOV = SqlMOV & " Where M.IDLOJA = " & mvarIDLOJA & vbNewLine
   SqlMOV = SqlMOV & " And M.IDPROD = " & mvarIDPROD
   
   sTbRow = ""
   sTbRow = sTbRow & "Select Row_Number() Over (Order By DTMOV, IDMOV ) [NLAN]"
   sTbRow = sTbRow & ", *" & vbNewLine
   sTbRow = sTbRow & " From (" & SqlMOV & ") TbRow" & vbNewLine
    
  
   sTbRow2 = ""
   sTbRow2 = "Select Row_Number() Over (Order By DTMOV, IDMOV ) [NLAN]" & vbNewLine
   sTbRow2 = sTbRow2 & ", *" & vbNewLine
   sTbRow2 = sTbRow2 & " From (" & vbNewLine
   sTbRow2 = sTbRow2 & "Select M.IDLOJA , M.IDEST, M.IDPROD, M.DTMOV, M.IDMOV" & vbNewLine
   sTbRow2 = sTbRow2 & " , [CREDITO]   = Case WHEN M.QTDITEM>0 THEN M.QTDITEM    ELSE 0 End" & vbNewLine
   sTbRow2 = sTbRow2 & " , [DEBITO]     = Case WHEN M.QTDITEM<0 THEN -1*M.QTDITEM ELSE 0 End" & vbNewLine
   sTbRow2 = sTbRow2 & " From SMOVEST M " & vbNewLine
   sTbRow2 = sTbRow2 & " Where M.IDLOJA = " & mvarIDLOJA & vbNewLine
   sTbRow2 = sTbRow2 & " And M.IDPROD = " & mvarIDPROD
   sTbRow2 = sTbRow2 & ") TbRow" & vbNewLine
   
   
   SqlSaldo = ""
   SqlSaldo = SqlSaldo & " Select Sum(CREDITO)- Sum(DEBITO)"
   SqlSaldo = SqlSaldo & " From (" & sTbRow2 & ") TbSum"
   SqlSaldo = SqlSaldo & " Where TbSum.NLAN <= M.NLAN"
   
   Sql = "Select M.*" & vbNewLine
   If bLimpaGrd Then
      Sql = Sql & ", Null [Núm], Null [Data], Null [Fornecedor/Cliente]" & vbNewLine
      Sql = Sql & ", Null [Entrada], Null [Saída], Null [Saldo]" & vbNewLine
      Sql = Sql & " From SMOVEST M " & vbNewLine
      Sql = Sql & " Where M.IDLOJA<0" & vbNewLine
   Else
      Sql = Sql & ", M.NUMDOC [Núm], Convert(Varchar, M.DTMOV, 103) [Data], M.FAVORECIDO [Fornecedor/Cliente]" & vbNewLine
      Sql = Sql & ", [Entrada]= Case M.CREDITO WHEN 0 THEN Null ELSE M.CREDITO End" & vbNewLine
      Sql = Sql & ", [Saída]  = Case M.DEBITO   WHEN 0 THEN Null ELSE M.DEBITO   End" & vbNewLine
      Sql = Sql & ", (" & SqlSaldo & ") [Saldo]" & vbNewLine
      If mvarSys.GetParam("FINNMFAVORECIDO") = 1 Then
      
      End If
      Sql = Sql & " From (" & sTbRow & ") M " & vbNewLine
      Sql = Sql & " Where M.NLAN > 0" & vbNewLine
      Sql = Sql & " And M.IDPROD = " & mvarIDPROD
   End If
   QueryGrid = Sql
End Function
Private Function RegLan(sCampo As String, Optional pGrid As Object) As ReportRecordItem
   Dim nIndex As Integer
   Dim oRecI  As ReportRecordItem
   Dim oRec As ReportRecord
   
   If pGrid Is Nothing Then Set pGrid = mvarMe.GrdMov
   If pGrid.SelectedRows.Count <> 1 Then
      Set oRec = New ReportRecord
      Set oRecI = oRec.AddItem("")
      oRecI.Value = ""
      Set RegLan = oRecI
      Exit Function
   End If
   sCampo = UCase(sCampo)
   
   If xVal(GetTag(pGrid, "COUNT", 0)) = 0 Then
      Call SetTag(pGrid, "COUNT", pGrid.Columns.Count)
      For nIndex = 0 To pGrid.Columns.Count - 1
         Call SetTag(pGrid, UCase(mvarMe.GrdMov.Columns(nIndex).Caption), nIndex)
      Next
   End If
   nIndex = xVal(GetTag(pGrid, sCampo, 0))
   
   Set RegLan = pGrid.SelectedRows(0).Record(nIndex)
End Function
Private Sub LimpaGrd()
   Dim Sql As String
   
   On Error GoTo TrataErro
   
   Sql = QueryGrid(True)
   
   Set RsLan = Nothing
   Call mvarSys.xDb.AbreTabela(Sql, RsLan)
   Call FillRCFromRS(RsLan, mvarMe.GrdMov)
   mvarMe.GrdMov.Records.DeleteAll
   
   Call EscondeCampos
   Exit Sub
TrataErro:
   ShowError
End Sub
Private Sub EscondeCampos()
   Dim i As Integer
   
   Set iCampo = New Collection
   
   If RsLan Is Nothing Then Exit Sub
   If RsLan.State = 0 Then Exit Sub
   
   
   Call SetTag(mvarMe.GrdMov, "COUNT", RsLan.Fields.Count)
   For i = 0 To RsLan.Fields.Count - 1
      If ExisteItem(iCampo, RsLan.Fields(i).Name) Then
         mvarMe.GrdMov.Columns(i).Visible = False
      Else
         iCampo.Add i, RsLan.Fields(i).Name
         mvarMe.GrdMov.Columns(i).Visible = InArray(RsLan.Fields(i).Name, Array("Núm", "Data", "Fornecedor/Cliente", "Entrada", "Saída", "Saldo"))
         
         Call SetTag(mvarMe.GrdMov, UCase(RsLan.Fields(i).Name), i)
      End If
   Next
End Sub
Private Sub DefineGrdIcones()
   Dim i As Integer
   With mvarMe.GrdMov
      .SetImageList mvarMe.ImageList1
      
      For i = 0 To .Records.Count - 1
         .Records(i).Item(1).Icon = 0
      Next
   End With
End Sub
Private Sub DefineGrdFiltro()
   With mvarMe.GrdMov
      If mvarMe.txtFiltrar.Text = "" Or mvarMe.txtFiltrar.Text = StrPesquisa Then
         .FilterText = ""
      Else
         .FilterText = mvarMe.txtFiltrar.Text
      End If
   End With
End Sub
Private Sub DefineGrdAlinha()
   Dim vCol As ReportColumn
   With mvarMe.GrdMov
      For Each vCol In .Columns
         vCol.HeaderAlignment = xtpAlignmentCenter
         vCol.Sortable = False
      Next
      
      .Columns(iCampo("Data")).Alignment = xtpAlignmentCenter
      .Columns(iCampo("Saída")).Alignment = xtpAlignmentRight
      .Columns(iCampo("Entrada")).Alignment = xtpAlignmentRight
      .Columns(iCampo("Saldo")).Alignment = xtpAlignmentRight
      
      .AutoColumnSizing = False
      .Columns(iCampo("Núm")).Width = 60
      .Columns(iCampo("Data")).Width = 90
      .Columns(iCampo("Saída")).Width = 100
      .Columns(iCampo("Entrada")).Width = 100
      .Columns(iCampo("Saldo")).Width = 100
      
      .Columns(iCampo("Fornecedor/Cliente")).Width = (.Width / Screen.TwipsPerPixelX) - 502
      '.Populate
   End With
End Sub
Private Function RecalculaSaldo() As Boolean
   Dim Sql As String
   
End Function
Public Sub L_FillRCFromRS(ByRef pRecordSet As Object, ByRef pReportControl As Object, Optional bDoEvents As Boolean = False, Optional bDeleteColumn As Boolean = False)  ' XtremeReportControl)
   Dim xField     As Object   ' ADODB.Field
   Dim xRecord    As XtremeReportControl.ReportRecord
   Dim xItem      As XtremeReportControl.ReportRecordItem
   Dim nIndex     As Integer
   Dim nMouse     As Integer
   Dim nType      As Integer
   
   nMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
      
   pReportControl.Records.DeleteAll
   
   If Not bDeleteColumn Then
      bDeleteColumn = (pReportControl.Columns.Count <> pRecordSet.Fields.Count)
   End If
   
   If bDeleteColumn Then
      pReportControl.Columns.DeleteAll
      pReportControl.Redraw
      For Each xField In pRecordSet.Fields
         pReportControl.Columns.Add nIndex, xField.Name, 100, True
         nIndex = nIndex + 1
      Next xField
   End If
   
   Do While Not pRecordSet.EOF
      Set xRecord = pReportControl.Records.Add()
      For nIndex = 0 To pRecordSet.Fields.Count - 1
         Set xItem = xRecord.AddItem("")
         With xItem
            If IsNull(pRecordSet.Collect(nIndex)) Then
               .Value = Empty
            Else
               nType = pRecordSet.Fields(nIndex).Type
               If nType = adCurrency Or nType = adDecimal Or nType = adDouble Or nType = adNumeric Then
                  .Value = pRecordSet.Collect(nIndex)
                  If .Value < 0 Then
                     .Caption = "(" & Mid(ValBr(pRecordSet.Collect(nIndex), 0), 2) & ")"
                     .ForeColor = vbRed
                  Else
                     .Caption = ValBr(pRecordSet.Collect(nIndex), 0)
                  End If
               Else
                  .Value = pRecordSet.Collect(nIndex)
               End If
            End If
         End With
      Next
      If bDoEvents Then DoEvents
      pRecordSet.MoveNext
   Loop
   pReportControl.Populate
   Screen.MousePointer = nMouse
End Sub
Private Sub mvarPane_CmbPeriodoClick()
   If Val(GetTag(mvarPane.CmbPeriodo, "Index", 1)) <> mvarPane.CmbPeriodo.ListIndex Then
      Call PopulaGrdLan
      Call SetTag(mvarPane.CmbPeriodo, "Index", mvarPane.CmbPeriodo.ListIndex)
   End If
End Sub
