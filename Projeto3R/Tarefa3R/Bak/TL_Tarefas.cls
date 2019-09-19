VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Tarefas"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const StrPesquisa = "Pesquisar Extrato"
Private WithEvents mvarMe     As FrmTarefas
Attribute mvarMe.VB_VarHelpID = -1
'Private WithEvents mvarPane   As FrmPaneFinanc

Private RsLan As Recordset  'Object 'Object
Dim iCampo As Collection
Private bNovo As Boolean

Const TaskCC_Mov = 1
Const TaskCC_Det = 2
Const TaskCC_His = 3
Const ID_TASKITEM_NEWFOLDER = 4
Const ID_TASKITEM_PUBLISH = 5
Const ID_TASKITEM_SHARE = 6
Const ID_TASKITEM_MYCOMPUTER = 7
Const ID_TASKITEM_MYDOCUMENTS = 8
Const ID_TASKITEM_SHAREDDOCUMENTS = 9
Const ID_TASKITEM_MYNETWORKPLACES = 10

'Const FCONTROL = 8

Private mvarSys      As Object
Private mvarIDLOJA   As Integer
Private mvarIDCONTA  As Integer
Private mvarIDLAN    As Long
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
   mvarIDCONTA = mvarSys.Propriedades("IDCONTA")
End Property
Public Property Get Sys() As Object
   Set Sys = mvarSys
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDCONTA(ByVal vData As Integer)
   mvarIDCONTA = vData
End Property
Public Property Get IDCONTA() As Integer
   IDCONTA = mvarIDCONTA
End Property
Public Property Get IDLAN() As Long
   mvarIDLAN = mvarIDLAN
End Property
Public Sub Show()
   If mvarIDCONTA = 0 Then
      Call ExibirAviso("Não existe conta corrente definida.")
   End If
   
   '* Carregar e Acoplar Form no MDI
   Load mvarMe
   Call AcoplarForm(mvarMe, 2, mvarSys)
   
   '* Definir foco
   With mvarMe.GrdMovCC
      If .Visible And .Enabled Then
         .SetFocus
      End If
   End With
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmTarefas
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_Activate()
'   Call PopulaTela
   If GetTag(mvarMe, "LOAD", 1) = 1 Then
      Call SetTag(mvarMe, "LOAD", 0)
      Call mvarMe_TabContaSelectedChanged(mvarMe.TabConta.Item(mvarMe.TabConta.SelectedItem))
      Call mvarMe_GrdMovCCSelectionChanged
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmbCategoriaClick()
   Call MontaSubCategoria
End Sub
Private Sub mvarMe_CmbCategoriaLostFocus()
   Call SalvarCategoria
   If mvarMe.CmbCategoria.ListIndex < 0 Then
      If LocalizarCombo(mvarMe.CmbCategoria, mvarMe.CmbCategoria.Text, False) >= 0 Then
         Call LocalizarCombo(mvarMe.CmbCategoria, mvarMe.CmbCategoria.Text, True)
      End If
   End If
End Sub
Private Function SalvarCategoria(Optional pIDPAI As Long = 0) As Boolean
   Dim bAchou As Boolean
   Dim sDSC   As String
   Dim iTab   As Integer
   Dim TbDesp As Object
   Dim oCombo As XtremeSuiteControls.ComboBox
   Dim Sql     As String
   Dim sNewCod As String
   
   iTab = mvarMe.TabConta.SelectedItem
   
'   sNewCod = CStr(IIf(iTab = 0, 1, 2)) & "." & StrZero(mvarMe.CmbCategoria.ListCount + 1, 2)
'   Sql = "Select Max(CODDESP) [COD]"
'   Sql = Sql & " From FDESPESA"
'   Sql = Sql & " Where CODDESP like '" & CStr(IIf(iTab = 0, 1, 2)) & ".%"
'   Sql = Sql & " And IDPAI = 0 "
'   If mvarSys.xDb.Abretabela(Sql) Then
'      sNewCod = xVal(Mid(mvarSys.xDb.RsAux("COD"), 3)) + 1
'      sNewCod = CStr(IIf(iTab = 0, 1, 2)) & "." & StrZero(sNewCod + 1, 2)
'   End If
   
   If pIDPAI = 0 Then
      Set oCombo = mvarMe.CmbCategoria
   Else
      Set oCombo = mvarMe.CmbSubCategoria
      
'      Sql = "Select Max(CODDESP)"
'      Sql = Sql & " From FDESPESA"
'      Sql = Sql & " Where CODDESP like '" & CStr(IIf(iTab = 0, 1, 2)) & ".%"
'      Sql = Sql & " And IDPAI <> 0 "
'      If mvarSys.xDb.Abretabela(Sql) Then
'         sNewCod = xVal(Mid(mvarSys.xDb.RsAux("COD"), 3)) + 1
'         sNewCod = CStr(IIf(iTab = 0, 1, 2)) & "." & StrZero(sNewCod + 1, 2)
'      Else
'         sNewCod = CStr(IIf(iTab = 0, 1, 2)) & "." & StrZero(mvarMe.CmbCategoria.ListCount + 1, 2) & "." & StrZero(mvarMe.CmbSubCategoria.ListCount + 1, 2)
'      End If
   End If
   sNewCod = CStr(IIf(iTab = 0, 1, 2)) & "." & StrZero(oCombo.ListCount + 1, 2) & "." & StrZero(mvarMe.CmbSubCategoria.ListCount + 1, 2)
   sDSC = Trim(oCombo.Text)
   bAchou = (LocalizarCombo(oCombo, sDSC, False) >= 0)
   
   If Not bAchou And sDSC <> "" Then
      If ExibirPergunta("Deseja cadastrar a categoria '" & sDSC & "'") = vbNo Then
         oCombo.Text = ""
         oCombo.ListIndex = -1
      Else
         Set TbDesp = CriarObjeto("BANCO_3R.TB_FDESPESA")
         With TbDesp
            .xDb = mvarSys.xDb
            .ATIVO = 1
            .CODDESP = sNewCod
            .DSCDESP = sDSC
            .IDGRPDESP = 0
            .IDPAI = pIDPAI
            .OCULTO = 0
            .TPDESP = IIf(iTab = 0, 1, 2)
            If .Salvar Then
               SalvarCategoria = True
               oCombo.AddItem sDSC
               oCombo.ItemData(oCombo.NewIndex) = .IDDESP
               oCombo.ListIndex = oCombo.NewIndex
            End If
            'Call LocalizarCombo(oCombo, sDSC, True)
         End With
         Set TbDesp = Nothing
      End If
   Else
      Call LocalizarCombo(oCombo, sDSC, True)
   End If
End Function
Private Sub mvarMe_CmbDeLostFocus()
   Call PopulaUltTransf
End Sub

Private Sub mvarMe_CmbDTBAIXAChange()
   mvarMe.TxtDTBAIXA.Text = Format(mvarMe.CmbDTBAIXA.Value, "dd/mm/yyyy")
End Sub

Private Sub mvarMe_CmbFavorecidoLostFocus()
   If bNovo Then
      If LocalizarCombo(mvarMe.CmbFavorecido, mvarMe.CmbFavorecido.Text, False) >= 0 Then
         Call LocalizarCombo(mvarMe.CmbFavorecido, mvarMe.CmbFavorecido.Text)
         Call PopulaUltLan(mvarMe.CmbFavorecido.ItemData(mvarMe.CmbFavorecido.ListIndex))
      End If
   End If
End Sub
Private Sub PopulaUltTransf()
   Dim Sql        As String
   Dim oRs        As Object
   Dim bAchou     As Boolean
   Dim sData      As String
   Dim nIDCONTA   As Long
   
   If mvarMe.CmbPara.ListIndex < 0 Then Exit Sub
   If mvarMe.TabConta.SelectedItem <> 1 Then Exit Sub
   If xVal(mvarMe.TxtValor.Text) <> 0 And Trim(mvarMe.CmbFavorecido.Text) <> "" Then Exit Sub
         
   sData = Format(mvarMe.TxtDTBAIXA.Text, "dd/mm/yyyy")
   If Not IsDate(sData) Then
      sData = Format(mvarSys.xDb.sysdate, "dd/mm/yyyy")
   End If
   sData = sData & " 23:59"
   nIDCONTA = mvarMe.CmbPara.ItemData(mvarMe.CmbPara.ListIndex)
      
   Sql = "Select Top 1 IDFAVORECIDO, VALOR, HISTORICO, Max(DTBAIXA) [Data]"
   Sql = Sql & " From FLAN"
   Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
   Sql = Sql & " And IDCONTA=" & SqlNum(nIDCONTA)
   Sql = Sql & " And TPLAN='C'"
   Sql = Sql & " AND TPTRANSA='T'"
   Sql = Sql & " And DTBAIXA <= " & SqlStr(sData)
   Sql = Sql & " Group By IDFAVORECIDO, VALOR, HISTORICO"
   Sql = Sql & " Order By Data Desc"
   
   bAchou = mvarSys.xDb.AbreTabela(Sql, oRs)
   
   If Not bAchou Then
      Sql = "Select Top 1 IDFAVORECIDO, VALOR, HISTORICO, Min(DTBAIXA) [Data]"
      Sql = Sql & " From FLAN"
      Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
      Sql = Sql & " And IDCONTA=" & SqlNum(nIDCONTA)
      Sql = Sql & " And TPLAN='C'"
      Sql = Sql & " AND TPTRANSA='T'"
      Sql = Sql & " And DTBAIXA > " & SqlStr(sData)
      Sql = Sql & " Group By IDFAVORECIDO, VALOR, HISTORICO"
      Sql = Sql & " Order By Data"
   
      bAchou = mvarSys.xDb.AbreTabela(Sql, oRs)
   End If
   If bAchou Then
      If Trim(mvarMe.CmbFavorecido.Text) = "" And xVal(oRs("IDFAVORECIDO") & "") > 0 Then
         mvarMe.CmbFavorecido.ListIndex = -1
         Call LocalizarCombo(mvarMe.CmbFavorecido, oRs("IDFAVORECIDO") & "", True, True)
      End If
      If xVal(mvarMe.TxtValor.Text) = 0 Then
         mvarMe.TxtValor.Text = ValBr(xVal(oRs("VALOR") & ""))
      End If
      If Trim(mvarMe.TxtObs.Text) <> "" Then
         mvarMe.TxtObs.Text = oRs("HISTORICO") & ""
      End If
   End If
End Sub
Private Sub PopulaUltLan(pIDCONTATO As Long)
   Dim Sql As String
   Dim oRs As Object
   Dim bAchou As Boolean
   Dim sData  As String
   
   If pIDCONTATO <= 0 Then
      Exit Sub
   End If
   
   sData = Format(mvarMe.TxtDTBAIXA.Text, "dd/mm/yyyy")
   If Not IsDate(sData) Then
      sData = Format(mvarSys.xDb.sysdate, "dd/mm/yyyy")
   End If
   sData = sData & " 23:59"
      
   Sql = "Select Top 1 IDDESP, IDSUBDESP, VALOR, HISTORICO, Max(DTBAIXA) [Data]"
   Sql = Sql & " From FLAN"
   Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
   Sql = Sql & " And IDCONTA=" & SqlNum(mvarIDCONTA)
   Sql = Sql & " And IDFAVORECIDO=" & SqlNum(pIDCONTATO)
   Sql = Sql & " And DTBAIXA <=" & SqlStr(sData)
   Sql = Sql & " Group By IDDESP, IDSUBDESP, VALOR, HISTORICO"
   Sql = Sql & " Order By Data Desc"

   bAchou = mvarSys.xDb.AbreTabela(Sql, oRs)
   If Not bAchou Then
      Sql = "Select Top 1 IDDESP, IDSUBDESP, VALOR, HISTORICO, Min(DTBAIXA) [Data]"
      Sql = Sql & " From FLAN"
      Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
      Sql = Sql & " And IDCONTA=" & SqlNum(mvarIDCONTA)
      Sql = Sql & " And IDFAVORECIDO=" & SqlNum(pIDCONTATO)
      Sql = Sql & " And DTBAIXA >" & SqlStr(sData)
      Sql = Sql & " Group By IDDESP, IDSUBDESP, VALOR, HISTORICO"
      Sql = Sql & " Order By Data"
      
      bAchou = mvarSys.xDb.AbreTabela(Sql, oRs)
   End If
   If bAchou Then
      If Trim(mvarMe.CmbCategoria.Text) = "" And xVal(oRs("IDDESP") & "") > 0 Then
         mvarMe.CmbCategoria.ListIndex = -1
         Call LocalizarCombo(mvarMe.CmbCategoria, oRs("IDDESP") & "", True, True)
      End If
      If Trim(mvarMe.CmbSubCategoria.Text) = "" And xVal(oRs("IDSUBDESP") & "") > 0 Then
         mvarMe.CmbSubCategoria.ListIndex = -1
         Call LocalizarCombo(mvarMe.CmbSubCategoria, oRs("IDSUBDESP") & "", True, True)
      End If
      If xVal(mvarMe.TxtValor.Text) = 0 Then
         mvarMe.TxtValor.Text = ValBr(xVal(oRs("VALOR") & ""))
      End If
      If Trim(mvarMe.TxtObs.Text) <> "" Then
         mvarMe.TxtObs.Text = oRs("HISTORICO") & ""
      End If
   End If
End Sub
Private Sub mvarMe_CmbParaLostFocus()
   Call PopulaUltTransf
End Sub
Private Sub mvarMe_CmbSubCategoriaLostFocus()
   If mvarMe.CmbCategoria.ListIndex >= 0 Then
      Call SalvarCategoria(mvarMe.CmbCategoria.ItemData(mvarMe.CmbCategoria.ListIndex))
   End If
End Sub
Private Sub mvarMe_CmdCancelarClick()
   Call HabilitaTab(False)
   Call PopulaTab
   mvarMe.GrdMovCC.SetFocus
End Sub
Private Sub mvarMe_CmdNovoClick()
   Dim sData   As String
   Dim nIndex As Long
   
   
   If mvarMe.GrdMovCC.SelectedRows.Count > 0 Then
      nIndex = mvarMe.GrdMovCC.SelectedRows(0).Index
      sData = Format(mvarMe.GrdMovCC.SelectedRows(0).Record.Item(iCampo("DTBAIXA")).Value, "dd/mm/yyyy")
   End If
   Call HabilitaTab(True)
   Call InserirGrdLinha
      
   If mvarMe.TabConta.SelectedItem = 1 Then
      Call SelContaTrans
   Else
      If mvarMe.CmbFavorecido.Visible And mvarMe.CmbFavorecido.Enabled Then
         mvarMe.CmbFavorecido.SetFocus
      End If
   End If
       
   If mvarMe.GrdMovCC.Rows.Count > nIndex Then
      mvarMe.GrdMovCC.Rows(nIndex).Selected = True
   End If
   If IsDate(sData) Then
      mvarMe.TxtDTBAIXA.Text = sData
      mvarMe.CmbDTBAIXA.Value = CDate(mvarMe.TxtDTBAIXA.Text)
   End If
   bNovo = True
End Sub
Private Sub SelContaTrans()
   DoEvents
   mvarMe.CmbDe.ListIndex = -1
   mvarMe.CmbDe.ListIndex = LocalizarCombo(mvarMe.CmbDe, CStr(mvarIDCONTA), False, True)
   If mvarMe.CmbDe.ListIndex >= 0 Then
      If mvarMe.CmbPara.ListCount = 2 Then
         mvarMe.CmbPara.ListIndex = 0
         If mvarMe.CmbPara.ItemData(0) = mvarMe.CmbDe.ItemData(mvarMe.CmbDe.ListIndex) Then
            mvarMe.CmbPara.ListIndex = 1
         End If
         If mvarMe.TxtDTBAIXA.Enabled And mvarMe.TxtDTBAIXA.Visible Then
            mvarMe.TxtDTBAIXA.SetFocus
         End If
      Else
         If mvarMe.CmbPara.Enabled And mvarMe.CmbPara.Visible Then
            mvarMe.CmbPara.SetFocus
         End If
      End If
   Else
      If mvarMe.CmbDe.Enabled And mvarMe.CmbDe.Visible Then
         mvarMe.CmbDe.SetFocus
      End If
   End If
End Sub
Private Sub mvarMe_CmdNumClick()
   Dim Ng As NG_Tarefa
   
   Set Ng = New NG_Tarefa
   Set Ng.Sys = mvarSys
   Call Ng.ExibirVenda(mvarMe.TxtNDOC.Text)
End Sub
Private Sub mvarMe_CmdSalvarClick()
   If mvarMe.ActiveControl Is mvarMe.TxtDTBAIXA Then
      Call mvarMe_TxtDTBAIXALostFocus
   End If
   Call SalvarLAN
End Sub
Private Sub SalvarLAN()
   Dim bOk        As Boolean
   Dim bIsDirt    As Boolean
   Dim TbFLAN     As Object
   Dim TbTransf   As Object
   Dim nIndex     As Long
   Dim nTopRow    As Long
   Dim Queries    As Collection
   Dim nCtPara    As Integer
   Dim nCtDe      As Integer
   Dim oRptRow    As ReportRow
   
   If Not ValidaCampos() Then Exit Sub
   If Not VerificaLAN(False, mvarIDLAN) Then Exit Sub
   
   Set Queries = New Collection
   
   nTopRow = mvarMe.GrdMovCC.TopRowIndex
   Call PopulaClasse(TbFLAN, Queries)
   
   With TbFLAN
      bIsDirt = .isDirt
      bOk = True
      If bIsDirt Then
         bOk = .Salvar
         Set Queries = New Collection
         
         If bOk Then
            If .TPTRANSA = "T" Then
               nCtDe = mvarMe.CmbDe.ItemData(mvarMe.CmbDe.ListIndex)
               nCtPara = mvarMe.CmbPara.ItemData(mvarMe.CmbPara.ListIndex)
               
               mvarIDLAN = .IDTRANSF
               mvarIDCONTA = IIf(.IDCONTA = nCtDe, nCtPara, nCtDe)
               Call PopulaClasse(TbTransf)
               mvarIDLAN = .IDLAN
               mvarIDCONTA = .IDCONTA
               
               TbTransf.IDTRANSF = .IDLAN
               If TbTransf.isDirt Then
                  bOk = TbTransf.Salvar
               End If
               If bOk Then
                  .IDTRANSF = TbTransf.IDLAN
                  If .isDirt And bOk Then
                     bOk = .Salvar
                  End If
               End If
            End If
         End If
      End If
         
      If .IDFAVORECIDO = 0 Then Call SalvarFavorecido(TbFLAN)

      If bOk And bIsDirt Then
         Call PopulaGrdLan
         RsLan.MoveFirst
         RsLan.Find "IDLAN=" & .IDLAN
         nIndex = RsLan.AbsolutePosition - 1
         
         If nIndex >= 0 Then
            With mvarMe.GrdMovCC
               Set oRptRow = .Rows.FindRow(mvarMe.GrdMovCC.Records(nIndex))
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
   Call ExibirResultado(mvarSys, bOk, IIf(bOk, 1, 2))
'   Salvar = bOk
End Sub
Private Function ValidaCampos() As Boolean
   On Error GoTo TrataErro
   
   If Not IsDate(mvarMe.TxtDTBAIXA.Text) Then
      Call ExibirAviso("Data inválida.")
      mvarMe.TxtDTBAIXA.SetFocus
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
   If mvarMe.TabConta.SelectedItem = 1 Then
      If Trim(mvarMe.CmbDe.Text) = "" Then
         Call ExibirAviso("Conta origem inválida.")
         mvarMe.CmbDe.SetFocus
         GoTo Saida
      ElseIf mvarMe.CmbDe.ListIndex < 0 Then
         Call LocalizarCombo(mvarMe.CmbDe, mvarMe.CmbDe.Text)
      End If
      If mvarMe.CmbDe.ListIndex < 0 Then
         Call ExibirAviso("Conta origem inválida.")
         mvarMe.CmbDe.SetFocus
         GoTo Saida
      End If
      
      If Trim(mvarMe.CmbPara.Text) = "" Then
         Call ExibirAviso("Conta destino inválida.")
         mvarMe.CmbPara.SetFocus
         GoTo Saida
      ElseIf mvarMe.CmbPara.ListIndex < 0 Then
         Call LocalizarCombo(mvarMe.CmbPara, mvarMe.CmbPara.Text)
      End If
      If mvarMe.CmbPara.ListIndex < 0 Then
         Call ExibirAviso("Conta destino inválida.")
         mvarMe.CmbPara.SetFocus
         GoTo Saida
      End If
      
      If mvarMe.CmbDe.ItemData(mvarMe.CmbDe.ListIndex) = mvarMe.CmbPara.ItemData(mvarMe.CmbPara.ListIndex) Then
         Call ExibirAviso("Conta destino não pode ser iguaa a conta origem.")
         mvarMe.CmbPara.SetFocus
         GoTo Saida
      End If
   End If
                     
   ValidaCampos = True
TrataErro:
Saida:
End Function
Private Sub SalvarFavorecido(Optional TbFLAN As Object)
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
         .ATIVO = 1
         If .DTCADASTRO = "" Then .DTCADASTRO = mvarSys.xDb.sysdate
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
   
   If Not TbFLAN Is Nothing Then
      If TbFLAN.IDFAVORECIDO = 0 And mvarMe.CmbFavorecido.ListIndex >= 0 Then
         nIDCont = mvarMe.CmbFavorecido.ItemData(mvarMe.CmbFavorecido.NewIndex)
         
         TbFLAN.IDFAVORECIDO = nIDCont
         'TbFLAN.TPFAVORECIDO = IIf(TbContato.PJ = 1, "F", "C")
         If TbFLAN.Salvar Then
            mvarMe.CmbFavorecido.ListIndex = -1
            Call LocalizarCombo(mvarMe.CmbFavorecido, CStr(nIDCont), True, True)
         End If
      End If
   End If
End Sub
Private Sub ExcluirLAN()
   Dim TbFLAN  As Object
   Dim TbTran  As Object
   Dim bOk     As Boolean
   Dim Queries As Collection
         
   If vbNo = ExibirPergunta("Deseja realmente excluir o lançamento.", , False) Then GoTo Saida
         
   Set Queries = New Collection
   If VerificaLAN(bDelete:=True) Then
      Set TbFLAN = CriarObjeto("BANCO_3R.TB_FLAN")
      Set TbFLAN.xDb = mvarSys.xDb
      
      bOk = True
      If TbFLAN.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCONTA:=mvarIDCONTA, Ch_IDLAN:=mvarIDLAN) Then
         Queries.Add TbFLAN.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCONTA:=mvarIDCONTA, Ch_IDLAN:=mvarIDLAN)
         
         If TbFLAN.IDTRANSF <> 0 Then
            Set TbTran = CriarObjeto("BANCO_3R.TB_FLAN")
            Set TbTran.xDb = mvarSys.xDb
            
            bOk = True
            If TbTran.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDLAN:=TbFLAN.IDTRANSF) Then
               Queries.Add TbTran.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCONTA:=TbTran.IDCONTA, Ch_IDLAN:=TbTran.IDTRANSF)
            End If
         End If
         
         bOk = mvarSys.xDb.Executa(Queries)
      End If
      Set TbFLAN = Nothing
      Set Queries = Nothing
      
      If bOk Then
         Call PopulaGrdLan
         Call PopulaTab
      End If
      Call ExibirResultado(mvarSys, True, IIf(bOk, 1, 2))
   End If
Saida:

End Sub
Private Function VerificaLAN(Optional bDelete As Boolean, Optional pIDLAN As Long) As Boolean
   Dim nIDVENDA   As Long
   Dim nValor     As Double
   Dim sNDOC      As String
   Dim sMsg       As String
   Dim bOk        As Boolean
   Dim bPossuiVenda As Boolean
   Dim bExisteVenda As Boolean
   Dim bAlteraValor As Boolean
   Dim bAlteraNDOC  As Boolean
   
   
   If mvarMe.GrdMovCC.SelectedRows(0).Index >= mvarMe.GrdMovCC.Records.Count - 1 Then
      If Not bDelete Then VerificaLAN = True
      Exit Function
   End If
   With mvarMe.GrdMovCC.SelectedRows(0).Record
      If pIDLAN = 0 And bDelete Then
         pIDLAN = .Item(iCampo("IDLAN")).Value
      End If
      If pIDLAN <> 0 Then
         nIDVENDA = xVal(.Item(iCampo("IDVENDA")).Value)
         nValor = xVal(.Item(iCampo("VALOR")).Value)
         sNDOC = .Item(iCampo("NDOC")).Value
      End If
      bPossuiVenda = (nIDVENDA <> 0)
      bAlteraValor = (nValor <> xVal(mvarMe.TxtValor.Text))
      bAlteraNDOC = (sNDOC <> mvarMe.TxtNDOC.Text)
   End With
      
   bOk = True
   If bDelete Then
      If bPossuiVenda Then
         bExisteVenda = mvarSys.xDb.AbreTabela("Select * From CVENDA Where IDLOJA=" & mvarIDLOJA & " And IDVENDA=" & nIDVENDA)
         If bExisteVenda Then
            bOk = False
            sMsg = "Lançamento associado a venda '" & nIDVENDA & "' e não pode ter seu valor alterado ou excluído."
            Call ExibirAviso(sMsg)
         Else
            bOk = False
            sMsg = "Lançamento associado a venda '" & nIDVENDA & "'." & vbNewLine
            sMsg = sMsg & "Deseja continuar?"
            bOk = (ExibirPergunta(sMsg, "Financeiro", pDefaultYes:=False) = vbYes)
            If bOk Then
               bOk = (mvarSys.Propriedades("SENHAMESTRE") = InputBoxPassword("Entre com a senha mestre."))
            End If
         End If
      End If
   Else
      If bPossuiVenda And (bAlteraValor Or bAlteraNDOC) Then
         bOk = False
         Call ExibirAviso("Lançamento associado a venda '" & nIDVENDA & "' e não pode ter seu valor alterado ou excluído.")
         mvarMe.TxtNDOC.Text = sNDOC
         mvarMe.TxtValor.Text = ValBr(nValor)
      End If
   End If
   VerificaLAN = bOk
End Function
Private Sub PopulaClasse(ByRef TbFLAN As Object, Optional ByRef Queries As Collection)
   Dim sDate   As String
   Dim Sql     As String
   Dim TbTran  As Object
   
   Set TbFLAN = CriarObjeto("BANCO_3R.TB_FLAN")
   With TbFLAN
      .xDb = mvarSys.xDb
      
      bNovo = Not .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCONTA:=mvarIDCONTA, Ch_IDLAN:=mvarIDLAN)
      If bNovo Then
         .CODSIT = ""
         .DTCADASTRO = mvarSys.xDb.sysdate(3)
      End If
      .IDLOJA = mvarIDLOJA
      .IDCONTA = mvarIDCONTA
      '.IDLAN = mvarIDLAN
      
      If Format(.DTBAIXA, "dd/mm/yyyy") <> mvarMe.TxtDTBAIXA.Text Then
         .DTBAIXA = mvarMe.TxtDTBAIXA.Text & " 23:59"
      End If
      .DTEMISSAO = .DTBAIXA
      .DTVENCIMENTO = .DTBAIXA
      .IDFAVORECIDO = 0
      If mvarMe.CmbFavorecido.ListIndex >= 0 Then
         .IDFAVORECIDO = mvarMe.CmbFavorecido.ItemData(mvarMe.CmbFavorecido.ListIndex)
      End If
      If .IDFAVORECIDO = 0 Then .IDFAVORECIDO = .IDFAVORECIDO
      .FAVORECIDO = mvarMe.CmbFavorecido.Text
      .HISTORICO = mvarMe.TxtObs.Text
      .IDDESP = 0
      .IDSUBDESP = 0
      If mvarMe.CmbCategoria.ListIndex >= 0 Then .IDDESP = mvarMe.CmbCategoria.ItemData(mvarMe.CmbCategoria.ListIndex)
      If mvarMe.CmbSubCategoria.ListIndex >= 0 Then .IDSUBDESP = mvarMe.CmbSubCategoria.ItemData(mvarMe.CmbSubCategoria.ListIndex)
      .NDOC = mvarMe.TxtNDOC.Text
      
      If xVal(mvarMe.TxtValor.Text) < 0 Then
         If .TPTRANSA = "C" Then mvarMe.TabConta.SelectedItem = 2
         If .TPTRANSA = "T" Then mvarMe.TabConta.SelectedItem = 1
         If .TPTRANSA = "D" Then mvarMe.TabConta.SelectedItem = 0
         If .TPTRANSA = "Q" Then mvarMe.TabConta.SelectedItem = 0
         
         mvarMe.TxtValor.Text = ValBr(-1 * xVal(mvarMe.TxtValor.Text))
      End If
      .Valor = xVal(mvarMe.TxtValor.Text)
      If mvarMe.TabConta.SelectedItem = 0 Then .TPTRANSA = "C"
      If mvarMe.TabConta.SelectedItem = 1 Then .TPTRANSA = "T"
      If mvarMe.TabConta.SelectedItem = 2 Then .TPTRANSA = "D"
      If mvarMe.TabConta.SelectedItem = 3 Then .TPTRANSA = "Q"
      
      If Trim(.CODSIT) <> "" Then
         .DTCONTAB = mvarSys.xDb.sysdate
      End If
      '.TPFAVORECIDO="C" "F"
      '.IDPAI
      '.IDPGTO
      If .TPTRANSA = "T" Then
         .TPLAN = IIf(mvarIDCONTA = mvarMe.CmbDe.ItemData(mvarMe.CmbDe.ListIndex), "D", "C")
      Else
         .TPLAN = IIf(mvarMe.TabConta.SelectedItem = 0, "C", "D")
      End If
      '.IDVENDA
      '.NPARCELA
      '.TPARCELA
      If .TPDOC <> "CV" Then
         If .NDOC <> "" And .TPTRANSA <> "T" Then
            If .TPLAN = "C" Then
               .TPDOC = "CV"
            Else
               .TPDOC = "NF"
            End If
         End If
      End If
      
      .FLGDELETE = 0
   End With
End Sub
Private Sub CriarRegistro()
   Dim xRow As ReportRecord
   Dim xNewRow As ReportRecord
   Dim xItem As ReportRecordItem
   Dim i As Integer
   
   Set xRow = New ReportRecord
   For i = 0 To mvarMe.GrdMovCC.Columns.Count - 1
      Select Case mvarMe.GrdMovCC.Columns(i).Caption
         Case "IDLOJA":       xRow.AddItem mvarIDLOJA
         Case "IDCONTA":      xRow.AddItem mvarIDCONTA
         Case "TPLAN":        xRow.AddItem IIf(mvarMe.TabConta.SelectedItem = 0, "C", "D")
         Case "TPTRANSA"
            If mvarMe.TabConta.SelectedItem = 0 Then xRow.AddItem "C"
            If mvarMe.TabConta.SelectedItem = 1 Then xRow.AddItem "T"
            If mvarMe.TabConta.SelectedItem = 2 Then xRow.AddItem "D"
            If mvarMe.TabConta.SelectedItem = 3 Then xRow.AddItem "Q"
         Case "IDTRANSF":     xRow.AddItem IIf(mvarMe.TabConta.SelectedItem = 1, 0, Null)
         Case "HISTORICO":    xRow.AddItem mvarMe.TxtObs.Text
         Case "DTORDER":      xRow.AddItem mvarMe.TxtDTBAIXA.Text
         Case "DTEMISSAO":    xRow.AddItem mvarMe.TxtDTBAIXA.Text
         Case "DTVENCIMENTO": xRow.AddItem mvarMe.TxtDTBAIXA.Text
         Case "DTCONTAB":     xRow.AddItem Null
         Case "Núm":          xRow.AddItem mvarMe.TxtNDOC.Text
         Case "Data":         xRow.AddItem mvarMe.TxtDTBAIXA.Text
         Case "Favorecido":   xRow.AddItem mvarMe.CmbFavorecido.Text
         Case "C":            xRow.AddItem Null
         Case "Pagamento":    xRow.AddItem IIf(mvarMe.TabConta.SelectedItem = 0, Null, ValBr(mvarMe.TxtValor.Text))
         Case "Depósito":     xRow.AddItem IIf(mvarMe.TabConta.SelectedItem = 0, ValBr(mvarMe.TxtValor.Text), Null)
         Case "Saldo":        xRow.AddItem Null
         Case Else:
            xRow.AddItem Null
      End Select
   Next
   'Sql = "Select IDLOJA , IDCONTA, IDLAN, TPLAN, TPTRANSA, IDTRANSF, HISTORICO"
   'Sql = Sql & ", CONVERT(datetime, ISNULL(DTBAIXA, GETDATE()),105) AS DTORDER"
   'Sql = Sql & ", DTEMISSAO, DTVENCIMENTO, DTCONTAB"
   'Sql = Sql & ", NDOC [Núm], CONVERT(varchar, DTBAIXA,103) [Data], FAVORECIDO [Favorecido], CODSIT [C]"
   'Sql = Sql & ", [Pagamento]= Case TPLAN WHEN 'D' THEN VALOR ELSE NULL End"
   'Sql = Sql & ", [Depósito] = Case TPLAN WHEN 'C' THEN VALOR ELSE NULL End"
   'Sql = Sql & ", SALDO [Saldo]"
   i = 40
   mvarMe.GrdMovCC.Records.InsertAt i, xRow
   mvarMe.GrdMovCC.Populate
   mvarMe.GrdMovCC.Rows(i).EnsureVisible
   mvarMe.GrdMovCC.Rows(i).Selected = True
End Sub
Private Sub mvarMe_GrdMovCCBeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
   If Not iCampo Is Nothing Then
      'If Item.Index = iCampo("Depósito") Or Item.Index = iCampo("Pagamento") Or Item.Index = iCampo("Saldo") Then
      '   Item.Value = IIf(xVal(Item.Value & "") = 0, "", ValBr(Item.Value & ""))
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
Private Sub mvarMe_GrdMovCCRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   On Error Resume Next
   Call HabilitaTab(True)
   Select Case Item.Index
      Case iCampo("Núm"): mvarMe.TxtNDOC.SetFocus
      Case iCampo("Data"): mvarMe.TxtDTBAIXA.SetFocus
      Case iCampo("Favorecido"): mvarMe.CmbFavorecido.SetFocus
      Case iCampo("C"): mvarMe.TxtNDOC.SetFocus
      Case iCampo("Pagamento"): mvarMe.TxtValor.SetFocus
      Case iCampo("Depósito"): mvarMe.TxtValor.SetFocus
      Case iCampo("Saldo"): mvarMe.TxtValor.SetFocus
      Case Else:
         mvarMe.TxtNDOC.SetFocus
         If Trim(mvarMe.TxtNDOC.Text) = "" Then
            SendKeys "{Tab}"
         End If
   End Select
End Sub
Private Sub mvarMe_GrdMovCCSelectionChanged()
   
   Call LimpaTab
   Call HabilitaTab(False)
   Call PopulaTab
   If Not mvarMe.CmbFavorecido.Visible Then
      Call mvarMe_TabContaSelectedChanged(mvarMe.TabConta.Item(mvarMe.TabConta.SelectedItem))
   End If
   
   DoEvents
   mvarMe.GrdMovCC.SetFocus
End Sub
Private Sub LimpaTab()
   With mvarMe
      Call LocalizarCombo(.CmbDe, "")
      Call LocalizarCombo(.CmbPara, "")
      Call LocalizarCombo(.CmbFavorecido, "")
      Call LocalizarCombo(.CmbCategoria, "")
      Call LocalizarCombo(.CmbSubCategoria, "")
      .TxtObs.Text = ""
      .TxtNDOC.Text = ""
      .TxtDTBAIXA.Text = ""
      .TxtValor.Text = ""
   End With
End Sub
Private Sub HabilitaTab(pEnabled As Boolean)
   With mvarMe
'      If pEnabled Then
'         Call mvarMe_TabContaSelectedChanged(mvarMe.TabConta.Item(mvarMe.TabConta.SelectedItem))
'      End If
      If Not pEnabled Then
         .CmbDe.Text = ""
         .CmbPara.Text = ""
         .CmbFavorecido.Text = ""
         .CmbCategoria.Text = ""
         .CmbSubCategoria.Text = ""
      End If
      
      .CmbDe.Enabled = pEnabled
      .CmbPara.Enabled = pEnabled
      .CmbFavorecido.Enabled = pEnabled
      .CmbCategoria.Enabled = pEnabled
      .CmbSubCategoria.Enabled = pEnabled
      .TxtObs.Enabled = pEnabled
      .TxtNDOC.Enabled = pEnabled
      .TxtDTBAIXA.Enabled = pEnabled
      .CmbDTBAIXA.Enabled = pEnabled
      .TxtValor.Enabled = pEnabled
      
      mvarMe.CmdSalvar.Enabled = pEnabled
      mvarMe.CmdCancelar.Enabled = pEnabled
      mvarMe.CmdNovo.Enabled = Not pEnabled
   End With
End Sub
Private Sub PopulaTab()
   Dim nPag As Currency
   Dim nDep As Currency
   Dim oReg As ReportRecord
   
   With mvarMe
      If .GrdMovCC.SelectedRows.Count <> 1 Then Exit Sub

      Set oReg = .GrdMovCC.SelectedRows(0).Record

      mvarIDLAN = RegLan("IDLAN").Value
      'mvarIDCONTA = RegLan("IDCONTA").Value
      bNovo = (mvarIDLAN = 0)
      Select Case RegLan("TPTRANSA").Value
         Case "C": .TabConta.SelectedItem = 0
         Case "T": .TabConta.SelectedItem = 1
         Case "D": .TabConta.SelectedItem = 2
         Case "Q": .TabConta.SelectedItem = 3
         Case Else: If mvarIDLAN <> 0 Then .TabConta.SelectedItem = 0
      End Select
      
      If RegLan("TPTRANSA").Value = "T" Then
         
         Dim TbTransf   As Object
         Dim nIDLOJA    As Long
         Dim nIDCONTA_T As Long
         Dim nIDTRANSF  As Long
         
         nIDLOJA = RegLan("IDLOJA").Value
         nIDTRANSF = RegLan("IDTRANSF").Value
         
         nIDCONTA_T = xVal(GetTag(RegLan("TPLAN"), "IDCONTA_T", 0))
         If nIDCONTA_T = 0 Then
            Set TbTransf = CriarObjeto("BANCO_3R.TB_FLAN")
            With TbTransf
               Set .xDb = mvarSys.xDb
               If .Pesquisar(Ch_IDLOJA:=nIDLOJA, Ch_IDLAN:=nIDTRANSF) Then
                  nIDCONTA_T = .IDCONTA
                  Call SetTag(RegLan("TPLAN"), "IDCONTA_T", nIDCONTA_T)
               End If
            End With
            Set TbTransf = Nothing
         End If
         If RegLan("TPLAN").Value = "D" Then
            Call LocalizarCombo(.CmbDe, RegLan("IDCONTA").Value, True, True)
            Call LocalizarCombo(.CmbPara, CStr(nIDCONTA_T), True, True)
         Else
            Call LocalizarCombo(.CmbDe, CStr(nIDCONTA_T), True, True)
            Call LocalizarCombo(.CmbPara, RegLan("IDCONTA").Value, True, True)
         End If
      End If
      .TxtNDOC.Text = RegLan("Núm").Value
      .TxtDTBAIXA.Text = RegLan("Data").Value
      If IsDate(.TxtDTBAIXA.Text) Then
         .CmbDTBAIXA.Value = CDate(.TxtDTBAIXA.Text)
      End If
      
      nPag = xVal(RegLan("Pagamento").Value)
      nDep = xVal(RegLan("Depósito").Value)
      .TxtValor.Text = ValBr(nPag + nDep)
      
      .CmbFavorecido.ListIndex = -1
      .CmbCategoria.ListIndex = -1
      .CmbSubCategoria.ListIndex = -1
      
      .CmbFavorecido.ListIndex = LocalizarCombo(.CmbFavorecido, RegLan("IDFAVORECIDO").Value, False, True)
      .CmbFavorecido.Text = RegLan("Favorecido").Value

      .CmbCategoria.ListIndex = LocalizarCombo(.CmbCategoria, RegLan("IDDESP").Value, False, True)
      Call MontaSubCategoria
      .CmbSubCategoria.ListIndex = LocalizarCombo(.CmbSubCategoria, RegLan("IDSUBDESP").Value, False, True)
      .TxtObs.Text = RegLan("HISTORICO").Value
      
      If .TxtDTBAIXA.Text = "" Then .TxtValor.Text = ""
   End With
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
      If mvarMe.ActiveControl Is mvarMe.GrdMovCC Then
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
   Call MontaTela
   Call PopulaTela
   DoEvents
End Sub
Public Sub PopulaTela()
   Dim TbCC As Object
   Set TbCC = CriarObjeto("BANCO_3R.TB_FCCORRENTE")
   Set TbCC.xDb = mvarSys.xDb
   
   Call LimpaGrd
   If TbCC.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCONTA:=mvarIDCONTA) Then
      Call PopulaGrdLan
      'mvarMe_Rezise
      'Call mvarMe_TabContaSelectedChanged(mvarMe.TabConta.Item(mvarMe.TabConta.SelectedItem))
   
      mvarMe.SccConta.Caption = "Conta " & TbCC.DSCCONTA
   End If
End Sub
Private Sub MontaTela()
   'ReportControlGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   With mvarMe
      .GrpPg.ZOrder 1
      
      .txtFiltrar = StrPesquisa
      .SccConta.VisualTheme = xtpShortcutThemeOffice2003
      .SccConta.GradientHorizontal = False
      With .GrdMovCC
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
      
      '**********************************
      '* Conta Corrente
      '**********************************
      Dim MyTbConta As Object
      
      .CmbDe.AutoComplete = True
      .CmbPara.AutoComplete = True
      .CmbDe.Clear
      .CmbPara.Clear
      
      Set MyTbConta = CriarObjeto("BANCO_3R.TB_FCCORRENTE")
      With MyTbConta
         .xDb = mvarSys.xDb
         If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA) Then
            While Not .Rs.EOF
               .Popula
               mvarMe.CmbDe.AddItem .DSCCONTA
               mvarMe.CmbDe.ItemData(mvarMe.CmbDe.NewIndex) = .IDCONTA
               mvarMe.CmbPara.AddItem .DSCCONTA
               mvarMe.CmbPara.ItemData(mvarMe.CmbPara.NewIndex) = .IDCONTA
               .Rs.MoveNext
            Wend
         End If
      End With
      
      '**********************************
      '* Favorecido
      '**********************************
      Dim Sql As String
      Dim MyRs As Object
      
      .CmbFavorecido.AutoComplete = True
      .CmbFavorecido.Sorted = True
      .CmbFavorecido.Clear
      
      Sql = "Select Distinct C.IDCONTATO, C.NOME"
      Sql = Sql & " From OCONTATO C, FLAN F"
      Sql = Sql & " Where C.IDCONTATO=F.IDFAVORECIDO"
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
   
      '**********************************
      '* Categoria/Despesa
      '**********************************
      Dim MyTbDesp As Object

      .CmbCategoria.AutoComplete = True
      .CmbSubCategoria.AutoComplete = True
      .CmbCategoria.Clear
      .CmbSubCategoria.Clear

      Sql = "ATIVO=1"
      Sql = Sql & " And OCULTO=0"
      Sql = Sql & " And GRUPO=1"
      'Sql = Sql & " And TPCAT=1"
      
      Set MyTbDesp = CriarObjeto("BANCO_3R.TB_FDESPESA")
      With MyTbDesp
         .xDb = mvarSys.xDb
         If .Pesquisar(Ch_Where:=Sql) Then
            While Not .Rs.EOF
               .Popula
               mvarMe.CmbCategoria.AddItem .DSCDESP
               mvarMe.CmbCategoria.ItemData(mvarMe.CmbCategoria.NewIndex) = .IDDESP
               .Rs.MoveNext
            Wend
         End If
      End With
   End With
End Sub
Private Sub MontaCategoria()
   Dim Sql As String
   Dim MyTbDesp As Object

   With mvarMe
      .CmbCategoria.Clear
      
      Sql = ""
      Sql = Sql & "ATIVO=1"
      Sql = Sql & " And OCULTO=0 "
      Sql = Sql & " And TPDESP=" & IIf(mvarMe.TabConta.SelectedItem = 0, "1", "2")
      Sql = Sql & " And (IDPAI=0 or IDPAI Is Null)"
            
      Set MyTbDesp = CriarObjeto("BANCO_3R.TB_FDESPESA")
      With MyTbDesp
         .xDb = mvarSys.xDb
         If .Pesquisar(Ch_Where:=Sql) Then
            While Not .Rs.EOF
               .Popula
               mvarMe.CmbCategoria.AddItem .DSCDESP
               mvarMe.CmbCategoria.ItemData(mvarMe.CmbCategoria.NewIndex) = .IDDESP
               .Rs.MoveNext
            Wend
            If mvarMe.CmbCategoria.ListCount = 1 Then
               mvarMe.CmbCategoria.ListIndex = 0
            End If
         End If
      End With
   End With
End Sub
Private Sub MontaSubCategoria()
   Dim Sql As String
   Dim MyTbDesp As Object

   With mvarMe
      .CmbSubCategoria.Clear
      If .CmbCategoria.ListIndex < 0 Then Exit Sub
      
      Sql = ""
      Sql = Sql & "ATIVO=1"
      Sql = Sql & " And OCULTO=0 "
      Sql = Sql & " And TPDESP=" & IIf(mvarMe.TabConta.SelectedItem = 0, "1", "2")
      Sql = Sql & " And IDPAI=" & SqlNum(.CmbCategoria.ItemData(.CmbCategoria.ListIndex))
            
      Set MyTbDesp = CriarObjeto("BANCO_3R.TB_FDESPESA")
      With MyTbDesp
         .xDb = mvarSys.xDb
         If .Pesquisar(Ch_Where:=Sql) Then
            While Not .Rs.EOF
               .Popula
               mvarMe.CmbSubCategoria.AddItem .DSCDESP
               mvarMe.CmbSubCategoria.ItemData(mvarMe.CmbSubCategoria.NewIndex) = .IDDESP
               .Rs.MoveNext
            Wend
            If mvarMe.CmbSubCategoria.ListCount = 1 Then
               mvarMe.CmbSubCategoria.ListIndex = 0
            End If
         End If
      End With
   End With

End Sub
Private Sub mvarMe_Rezise()
   On Error Resume Next

   mvarMe.GrpPg.ZOrder 1
'   mvarMe.Refresh
   mvarMe.SccConta.Move 0, 50, mvarMe.ScaleWidth - 50, 420
   With mvarMe.txtFiltrar
      .Left = mvarMe.Width - 2800
      .Top = 90
   End With
   With mvarMe.imgLupa
      .Left = mvarMe.Width - 680
      .Top = 100
      .Height = mvarMe.txtFiltrar.Height - 25
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
      
   With mvarMe.GrdMovCC
      .Left = 0
      .Top = mvarMe.SccConta.Top + mvarMe.SccConta.Height + 30
      .Width = mvarMe.ScaleWidth - 50
      .Height = mvarMe.TabConta.Top - .Top - 30
   End With
   
   mvarMe.CmbDTBAIXA.Top = mvarMe.TxtDTBAIXA.Top
   mvarMe.CmbDTBAIXA.Height = mvarMe.TxtDTBAIXA.Height
   mvarMe.CmbDe.Width = mvarMe.GrpValor.Left - mvarMe.CmbDe.Left - 320
   mvarMe.CmbPara.Width = mvarMe.GrpValor.Left - mvarMe.CmbPara.Left - 320
   mvarMe.CmbFavorecido.Width = mvarMe.GrpValor.Left - mvarMe.CmbFavorecido.Left - 320
   mvarMe.CmbCategoria.Width = (mvarMe.CmbDe.Width / 2) - 30
   mvarMe.CmbSubCategoria.Left = mvarMe.CmbCategoria.Left + mvarMe.CmbCategoria.Width + 60
   mvarMe.CmbSubCategoria.Width = (mvarMe.CmbDe.Width / 2) - 30
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
      
   'DoEvents
   mvarMe.GrpPg.ZOrder 0
   'mvarMe.Refresh
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
         
      .LblDe.Visible = False
      .LblPara.Visible = False
      .LblFavorecido.Visible = False
      .LblCategoria.Visible = False
      .LblObs.Visible = False
      
      .LblFavorecido.Caption = "Pagar a:"
      .LblNum.TabIndex = 6
      .CmdNum.TabIndex = 7
      .TxtNDOC.TabIndex = 8
      
      Select Case Item.Index
         Case 0, 2, 3 '* Depóstio
            If Item.Index = 0 Then .LblFavorecido.Caption = "De:"
            '.LblDe.Top = Lin1
            '.LblPara.Top = Lin1
            .LblFavorecido.Top = Lin2
            .LblCategoria.Top = Lin3
            .LblObs.Top = Lin4
                  
            .LblFavorecido.TabIndex = 9
            .CmbFavorecido.TabIndex = 10
            .LblDTBAIXA.TabIndex = 11
            .TxtDTBAIXA.TabIndex = 12
            .LblCategoria.TabIndex = 13
            .CmbCategoria.TabIndex = 14
            .CmbSubCategoria.TabIndex = 15
            .LblVALOR.TabIndex = 16
            .TxtValor.TabIndex = 17
            .LblObs.TabIndex = 18
            .TxtObs.TabIndex = 19
            
                  
         Case 1 '* Transferência
            .LblDe.Top = Lin2
            .LblPara.Top = Lin3
            .LblFavorecido.Top = Lin4
            '.LblCategoria.Top = Lin4
            .LblObs.Top = Lin5
            
            .LblDe.TabIndex = 9
            .CmbDe.TabIndex = 10
            .LblDTBAIXA.TabIndex = 11
            .TxtDTBAIXA.TabIndex = 12
            .LblPara.TabIndex = 13
            .CmbPara.TabIndex = 14
            .LblVALOR.TabIndex = 15
            .TxtValor.TabIndex = 16
            .LblFavorecido.TabIndex = 17
            .CmbFavorecido.TabIndex = 18
            .LblObs.TabIndex = 19
            .TxtObs.TabIndex = 20
            
            If Not .CmdNovo.Enabled And .CmbDe.ListIndex < 0 Then
               Call SelContaTrans
            End If
            
         Case 2 '* Retirada
         Case 3 '* Cheque
      End Select
      Call MontaCategoria
      Call MontaSubCategoria

      .LblDe.Visible = InArray(Item.Index, Array(1))
      .LblPara.Visible = InArray(Item.Index, Array(1))
      .LblFavorecido.Visible = InArray(Item.Index, Array(0, 1, 2, 3))
      .LblCategoria.Visible = InArray(Item.Index, Array(0, 2, 3))
      .LblObs.Visible = InArray(Item.Index, Array(0, 1, 2, 3))
   
      .LblNum.Visible = InArray(Item.Index, Array(0, 1, 2, 3))
      .CmdNum.Visible = InArray(Item.Index, Array(0, 1, 2, 3))
      .LblDTBAIXA.Visible = InArray(Item.Index, Array(0, 1, 2, 3))
      .LblVALOR.Visible = InArray(Item.Index, Array(0, 1, 2, 3))
      '===================================================
      .CmbDe.Visible = .LblDe.Visible
      .CmbPara.Visible = .LblPara.Visible
      .CmbFavorecido.Visible = .LblFavorecido.Visible
      .CmbCategoria.Visible = .LblCategoria.Visible
      .CmbSubCategoria.Visible = .LblCategoria.Visible
      .TxtObs.Visible = .LblObs.Visible
      
      .TxtNDOC.Visible = (.LblNum.Visible Or .CmdNum.Visible)
      .CmbDTBAIXA.Visible = .LblDTBAIXA.Visible
      .TxtDTBAIXA.Visible = .LblDTBAIXA.Visible
      .TxtValor.Visible = .LblVALOR.Visible
      '===================================================
      
'      .LblDe.Top = Lin1
'      .LblPara.Top = Lin2
'      .LblFavorecido.Top = Lin2
'      .LblCategoria.Top = Lin3
      
      .CmbDe.Top = .LblDe.Top
      .CmbPara.Top = .LblPara.Top
      .CmbFavorecido.Top = .LblFavorecido.Top
      .CmbCategoria.Top = .LblCategoria.Top
      .CmbSubCategoria.Top = .LblCategoria.Top
      .TxtObs.Top = .LblObs.Top
      
      '.Refresh
      '.GrpPg.ZOrder 1
      .GrpPg.ZOrder 0
   End With
End Sub
Private Sub mvarMe_TxtDTBAIXALostFocus()
   mvarMe.TxtDTBAIXA.Text = FormatarData(mvarMe.TxtDTBAIXA.Text)
   Call PopulaUltTransf
End Sub
Private Sub mvarMe_txtFiltrarGotFocus()
   mvarMe.txtFiltrar.ForeColor = vbBlack
   
   If mvarMe.txtFiltrar.Text = StrPesquisa Then
      mvarMe.txtFiltrar.Text = ""
   End If
End Sub

Private Sub mvarMe_txtFiltrarKeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      mvarMe.GrdMovCC.FilterText = mvarMe.txtFiltrar.Text
      mvarMe.GrdMovCC.Populate
   End If
End Sub

Private Sub mvarMe_txtFiltrarLostFocus()
   
   If mvarMe.txtFiltrar.Text = "" Then mvarMe.txtFiltrar.Text = StrPesquisa
   mvarMe.txtFiltrar.ForeColor = 12632256
End Sub
Private Sub mvarMe_TxtNDOCLostFocus()
   If xVal(mvarMe.TxtNDOC.Text) > 0 Then
      mvarMe.TxtNDOC.Text = StrZero(mvarMe.TxtNDOC.Text, 6)
   End If
End Sub
Private Sub mvarMe_Unload(Cancel As Integer)
  On Error Resume Next
'   Unload FrmLembrete
End Sub
'=============================
'=============================
'Public Property Get PaneContact() As Object
'   Set PaneContact = mvarPane
'End Property
'Private Sub mvarPane_KeyUp(KeyCode As Integer, Shift As Integer)
'   Call mvarMe_KeyUp(KeyCode, Shift)
'End Sub
'Private Sub mvarPane_Load()
'   'MsgBox "mvarPaneLoad"
'   Call CreateTaskPanel
'End Sub
'Private Sub mvarPane_Resize()
'   mvarPane.SccConta.Move 0, 0, mvarPane.ScaleWidth, 420
'   mvarPane.SccConta2.Move 0, mvarPane.SccConta.Height, mvarPane.SccConta.Width, 285
'
'   mvarPane.Picture1.Move 0, mvarPane.SccConta2.Top + mvarPane.SccConta2.Height, mvarPane.ScaleWidth, mvarPane.ScaleHeight - mvarPane.SccConta.Height
'   mvarPane.wndTaskPanel.Move 0, 0, mvarPane.Picture1.ScaleWidth, mvarPane.Picture1.ScaleHeight
'
'End Sub
Private Sub PopulaGrdLan(Optional pWhere As String, Optional pIDLAN As Integer = 0)
   Dim Sql As String
   Dim i As Long
   Dim nTopRow As Long
   Dim nRowSel As Long
   Dim nTabSel As Integer
   Dim oRecord As ReportRecord
   Dim oRptRow As ReportRow
   
   Sql = QueryGrid
   Sql = Sql & IIf(pWhere = "", "", "AND " + pWhere)
   Sql = Sql & " Order By NLAN"
   
   Screen.MousePointer = vbHourglass
   nTabSel = mvarMe.TabConta.SelectedItem
   With mvarMe.GrdMovCC
      nTopRow = .TopRowIndex
      If .SelectedRows.Count <= 0 Then
         nRowSel = -1
      Else
         nRowSel = .SelectedRows(0).Index
         Set oRecord = .SelectedRows(0).Record
      End If
      
      .Visible = False
      If mvarSys.xDb.AbreTabela(Sql, RsLan) Then
      
         Call L_FillRCFromRS(RsLan, mvarMe.GrdMovCC)
         Call EscondeCampos
      
         If nRowSel < 0 Then
            nRowSel = RsLan.RecordCount
         End If
      Else
         Sql = QueryGrid(True)
         Call mvarSys.xDb.AbreTabela(Sql, RsLan)
         Call FillRCFromRS(RsLan, mvarMe.GrdMovCC)
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
      
      If pIDLAN <> 0 Then
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
   mvarIDLAN = 0
   With mvarMe.GrdMovCC
      If .Rows.Count > 0 Then
         '.Rows(.Rows.Count - 1).EnsureVisible
         .Rows(.Rows.Count - 1).Selected = True
         Call PopulaTab
      
         If .Rows(.Rows.Count - 1).Record(iCampo("IDLAN")).Value > 0 Then
            .InsertRecord .Rows.Count
            .Populate
            If .Rows.Count > 0 Then
               '.Rows(.Rows.Count - 1).EnsureVisible
               .Rows(.Rows.Count - 1).Selected = True
               Call PopulaTab
            End If
         End If
      Else
         .InsertRecord .Rows.Count
         .Populate
         If .Rows.Count > 0 Then
            '.Rows(.Rows.Count - 1).EnsureVisible
            .Rows(.Rows.Count - 1).Selected = True
            Call PopulaTab
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
   
   With mvarMe.GrdMovCC
      nSaldo = 0
      
      .ShowFooter = True
      For Each vCol In .Columns
         vCol.FooterText = " "
         vCol.DrawFooterDivider = False
         vCol.DrawFooterRowsVGrid = False
      Next

      '.Columns(iCampo("Núm")).FooterFont.Bold = True
      '.Columns(iCampo("Núm")).FooterAlignment = xtpAlignmentCenter
      '.Columns(iCampo("Núm")).FooterText = "Saldo Hoje :"
      
      '.Columns(iCampo("Data")).FooterFont.Bold = True
      '.Columns(iCampo("Data")).FooterAlignment = xtpAlignmentCenter
      '.Columns(iCampo("Data")).FooterText = "R$ " & ValBr(.Records(.Records.Count - 1).Item(iCampo("Saldo")).Value)
      
      .Columns(iCampo("Depósito")).FooterFont.Bold = True
      .Columns(iCampo("Depósito")).FooterAlignment = xtpAlignmentCenter
      .Columns(iCampo("Depósito")).FooterText = "Saldo Final :"
      
      .Columns(iCampo("Saldo")).FooterFont.Bold = True
      .Columns(iCampo("Saldo")).FooterAlignment = xtpAlignmentRight
      If .Records.Count > 0 Then
         .Columns(iCampo("Saldo")).FooterText = "R$ " & ValBr(.Records(.Records.Count - 1).Item(iCampo("Saldo")).Value)
      End If
   End With
End Sub
Public Function QueryGrid(Optional bLimpaGrd As Boolean) As String
   Dim Sql As String
   Dim SqlSaldo As String
   Dim SqlFLAN  As String
   Dim sTbRow As String
   
   SqlFLAN = ""
   SqlFLAN = SqlFLAN & "Select FL.IDLOJA , FL.IDCONTA, FL.IDLAN, FL.TPLAN, FL.TPTRANSA"
   SqlFLAN = SqlFLAN & ", FL.IDTRANSF, FL.HISTORICO, FL.DTBAIXA, FL.DTEMISSAO, FL.DTVENCIMENTO"
   SqlFLAN = SqlFLAN & ", FL.DTCONTAB, FL.NDOC, FL.IDFAVORECIDO, FL.TPFAVORECIDO, FL.FAVORECIDO [NOME]"
   SqlFLAN = SqlFLAN & ", FL.CODSIT, FL.IDVENDA, FL.IDPGTO, FL.VALOR, FL.IDDESP, FL.IDSUBDESP"
   SqlFLAN = SqlFLAN & ", [DEBITO]= Case FL.TPLAN WHEN 'D' THEN FL.VALOR ELSE 0 End"
   SqlFLAN = SqlFLAN & ", [CREDITO] = Case FL.TPLAN WHEN 'C' THEN FL.VALOR ELSE 0 End"
   SqlFLAN = SqlFLAN & " From FLAN FL"
   SqlFLAN = SqlFLAN & " Where FL.FLGDELETE = 0"
   SqlFLAN = SqlFLAN & " And Not FL.DTBAIXA is Null"
   SqlFLAN = SqlFLAN & " And FL.IDLOJA = " & SqlNum(mvarIDLOJA)
   SqlFLAN = SqlFLAN & " And FL.IDCONTA= " & SqlNum(mvarIDCONTA)
   
   sTbRow = ""
   'sTbRow = sTbRow & "Select Row_Number() Over (Order By DTBAIXA, DTEMISSAO, DTVENCIMENTO, IDLAN ) [NLAN]"
   sTbRow = sTbRow & "Select Row_Number() Over (Order By DTVENCIMENTO, DTEMISSAO, DTBAIXA, IDLAN ) [NLAN]"
   sTbRow = sTbRow & ", *"
   sTbRow = sTbRow & " From (" & SqlFLAN & ") TbRow"
   
   SqlSaldo = ""
   SqlSaldo = SqlSaldo & " Select Sum(CREDITO)- Sum(DEBITO)"
   SqlSaldo = SqlSaldo & " From (" & sTbRow & ") TbSum"
   SqlSaldo = SqlSaldo & " Where TbSum.NLAN <= F.NLAN"
   
'   Sql = ""
'   Sql = Sql & "Select F.IDLOJA , F.IDCONTA, F.IDLAN, F.TPLAN, F.TPTRANSA"
'   Sql = Sql & ", F.IDTRANSF, F.HISTORICO, F.DTBAIXA, F.DTEMISSAO, F.DTVENCIMENTO"
'   Sql = Sql & ", F.DTCONTAB, F.NDOC, F.IDFAVORECIDO, F.TPFAVORECIDO, F.FAVORECIDO"
'   Sql = Sql & ", F.CODSIT, F.IDVENDA, F.IDPGTO"
   
   Sql = "Select F.*"
   If bLimpaGrd Then
      Sql = "Select F.IDLOJA , F.IDCONTA, F.IDLAN, F.TPLAN, F.TPTRANSA"
      Sql = Sql & ", F.IDTRANSF, F.HISTORICO, F.DTBAIXA, F.DTEMISSAO, F.DTVENCIMENTO"
      Sql = Sql & ", F.DTCONTAB, F.NDOC, F.IDFAVORECIDO, F.TPFAVORECIDO, F.FAVORECIDO [NOME]"
      Sql = Sql & ", F.CODSIT, F.IDVENDA, F.IDPGTO, F.VALOR, F.IDDESP, F.IDSUBDESP"
      Sql = Sql & ", [DEBITO]= Case F.TPLAN WHEN 'D' THEN F.VALOR ELSE 0 End"
      Sql = Sql & ", [CREDITO] = Case F.TPLAN WHEN 'C' THEN F.VALOR ELSE 0 End"
   
      Sql = Sql & ", Null [Núm], Null [Data], Null [Favorecido], Null [C]"
      Sql = Sql & ", Null [Pagamento], Null [Depósito], Null [Saldo]"
      Sql = Sql & " From FLAN F "
      Sql = Sql & " Where F.IDLOJA<0"
   Else
      Sql = "Select F.*"
      'Sql = Sql & ", F.NDOC [Núm], Convert(varchar, F.DTBAIXA , 103) [Data], F.NOME [Favorecido], F.CODSIT [C]"
      Sql = Sql & ", F.NDOC [Núm], Convert(varchar, F.DTVENCIMENTO , 103) [Data], F.NOME [Favorecido], F.CODSIT [C]"
      Sql = Sql & ", [Pagamento] = Case F.DEBITO WHEN 0 THEN Null ELSE F.DEBITO End"
      Sql = Sql & ", [Depósito] = Case F.CREDITO WHEN 0 THEN Null ELSE F.CREDITO End"
      Sql = Sql & ", (" & SqlSaldo & ") [Saldo]"
      Sql = Sql & " From (" & sTbRow & ") F "
      Sql = Sql & " Where F.NLAN > 0"
   End If
   QueryGrid = Sql
End Function
Private Function RegLan(sCampo As String, Optional pGrid As Object) As ReportRecordItem
   Dim nIndex As Integer
   Dim oRecI  As ReportRecordItem
   Dim oRec As ReportRecord
   
   If pGrid Is Nothing Then Set pGrid = mvarMe.GrdMovCC
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
         Call SetTag(pGrid, UCase(mvarMe.GrdMovCC.Columns(nIndex).Caption), nIndex)
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
   Call FillRCFromRS(RsLan, mvarMe.GrdMovCC)
   mvarMe.GrdMovCC.Records.DeleteAll
   
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
   
   
   Call SetTag(mvarMe.GrdMovCC, "COUNT", RsLan.Fields.Count)
   For i = 0 To RsLan.Fields.Count - 1
      If ExisteItem(iCampo, RsLan.Fields(i).Name) Then
         mvarMe.GrdMovCC.Columns(i).Visible = False
      Else
         iCampo.Add i, RsLan.Fields(i).Name
         mvarMe.GrdMovCC.Columns(i).Visible = InArray(RsLan.Fields(i).Name, Array("Núm", "Data", "Favorecido", "C", "Pagamento", "Depósito", "Saldo"))
         
         Call SetTag(mvarMe.GrdMovCC, UCase(RsLan.Fields(i).Name), i)
      End If
   Next
End Sub
Private Sub DefineGrdIcones()
   Dim i As Integer
   With mvarMe.GrdMovCC
      .SetImageList mvarMe.ImageList1
      
      For i = 0 To .Records.Count - 1
         .Records(i).Item(1).Icon = 0
      Next
      '.Columns(1).Width = 20
      '.Columns(1).Caption = ""
      '.Columns(1).Icon = 1
   End With
End Sub
Private Sub DefineGrdFiltro()
   With mvarMe.GrdMovCC
      If mvarMe.txtFiltrar.Text = "" Or mvarMe.txtFiltrar.Text = StrPesquisa Then
         .FilterText = ""
      Else
         .FilterText = mvarMe.txtFiltrar.Text
      End If
   End With
End Sub
Private Sub DefineGrdAlinha()
   Dim vCol As ReportColumn
   With mvarMe.GrdMovCC
      For Each vCol In .Columns
         vCol.HeaderAlignment = xtpAlignmentCenter
         vCol.Sortable = False
      Next
      
      .Columns(iCampo("Data")).Alignment = xtpAlignmentCenter
      .Columns(iCampo("C")).Alignment = xtpAlignmentCenter
      .Columns(iCampo("Pagamento")).Alignment = xtpAlignmentRight
      .Columns(iCampo("Depósito")).Alignment = xtpAlignmentRight
      .Columns(iCampo("Saldo")).Alignment = xtpAlignmentRight
      
      .AutoColumnSizing = False
      .Columns(iCampo("Núm")).Width = 60
      .Columns(iCampo("Data")).Width = 90
      .Columns(iCampo("C")).Width = 30
      .Columns(iCampo("Pagamento")).Width = 100
      .Columns(iCampo("Depósito")).Width = 100
      .Columns(iCampo("Saldo")).Width = 100
      
      .Columns(iCampo("Favorecido")).Width = (.Width / Screen.TwipsPerPixelX) - 502
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
                     .Caption = "(" & Mid(ValBr(pRecordSet.Collect(nIndex)), 2) & ")"
                     .ForeColor = vbRed
                  Else
                     .Caption = ValBr(pRecordSet.Collect(nIndex))
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

