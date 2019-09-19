VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADFDESPESA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'** Palavras a serem substituidas
'** CADFDESPESA
'** FDESPESA
'** IDDESP
'** DSCDESP

Option Explicit
Const TbName = "FDESPESA"
Const IdName = "IDDESP"
Const Campo1 = "DSCDESP"
Const PergExc = "Confirma a exclusão da despesa/receita?"
Const Titulo = "Despesas/Receitas"

Private WithEvents mvarMe  As FrmCADFDESPESA
Attribute mvarMe.VB_VarHelpID = -1

Private mvarShortBar       As TL_ShortBar
Private mvarRptRec         As ReportRow
Private TbCad              As Object
Private mvarSys            As Object
Private mvarIDLOJA         As Integer
Private mvarID As Integer
Dim CollCOD As New Collection
Public Property Set TLCadastro(ByRef vData As Object)
   Dim Sql  As String
   Dim sAux As String
   With vData
      .Titulo = Titulo
      
      Sql = "Select " & IdName & " [Id.]"
      Sql = Sql & ", [Descrição]=Case IDPAI When 0 Then '' Else '     ' End + " & Campo1
      'Sql = Sql & ", " & Campo1 & " [Descrição]"
      Sql = Sql & ", [Tipo]=Case TPDESP When 2 Then 'Despesa' Else 'Receita' End"
      Sql = Sql & ", CODDESP [Código]"
      Sql = Sql & " From " & TbName
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      Sql = Sql & " Order By CODDESP"
      .Consulta = Sql
      
      sAux = ""
      sAux = sAux & "|CAMPO=" & IdName & "|CHAVE=S|GRUPO=0|VISIBLE=S:"
      .TAGCAMPOS = sAux
   End With
End Property
Public Property Let Id(ByVal vData As String)
   mvarID = vData
End Property
Public Property Get Id() As String
   Id = mvarID
End Property
Public Property Set RptRec(ByVal vData As ReportRow)
   Set mvarRptRec = vData
'   mvarIDLOJA = RptRec.Record.Item(0).Value
   If Not mvarRptRec Is Nothing Then mvarID = mvarRptRec.Record.Item(0).Value
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
   Set mvarMe = New FrmCADFDESPESA
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
   Set TbCad = CriarObjeto("BANCO_3R.TB_" & TbName)
   Set TbCad.Xdb = mvarSys.Xdb
End Sub
Private Sub PopulaTela()
   If TbCad Is Nothing Then
      Call InstanciaClasse
   End If
   
   With TbCad
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & mvarID) Then
         mvarMe.CmdExcluir.Enabled = True
         
         mvarMe.TxtID.Text = .IDDESP
         mvarMe.TxtCampo01.Text = .DSCDESP

      Else
         Call LimpaTela
      End If
   End With
End Sub
Public Function Excluir() As Boolean
   Dim Queries As Collection
   
   
   If mvarID <> 0 Then
      If VerificarDependencias() Then
         If ExibirPergunta(PergExc, "Atenção", False) = vbYes Then
            Call TbCad.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & xVal(mvarMe.TxtID))
            
            Set Queries = New Collection
            Call ExcluirDependencias(Queries)
            Queries.Add TbCad.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDDESP:=xVal(mvarMe.TxtID))
            
            If mvarSys.Xdb.Executa(Queries) Then
               Excluir = True
               Call ExibirResultado(mvarSys, True)
               
               'mvarIsDirt = True
               
               Call LimpaTela
               If Not mvarShortBar Is Nothing Then
                  mvarShortBar.TLCadastro.F_Refresh
               End If
               If mvarMe.TxtID.Visible And mvarMe.TxtCampo01.Enabled Then
                  mvarMe.TxtCampo01.SetFocus
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
   'Dim Tb01 As Object
   
   'Set Tb01 = Nothing
   'Set Tb01 = CriarObjeto("BANCO_3R.TB_OSALA_MAQUINA")
   'Set Tb01.XDb = mvarSys.XDb
            
   'Queries.Add Tb01.QryDelete(Ch_IDMAQUINA:=SqlNum(mvarMe.TxtID))
End Sub
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
      For Each n In mvarMe.Controls
         n.Text = ""
      Next
   End With
End Sub
Private Sub MontaTela()
   Dim MyTb As Object
   
   With mvarMe
      .CmbTPDESP.AddItem "Receita"
      .CmbTPDESP.AddItem "Despesa"
      .CmbTPDESP.ItemData(0) = 1
      .CmbTPDESP.ItemData(1) = 2
      .CmbTPDESP.ListIndex = 1
   End With
   
      
End Sub
Private Sub PopulaCmbIDPAI()
   Dim MyTb As Object
   Set MyTb = CriarObjeto("BANCO_3R.TB_FDESPESA")
   Set MyTb.Xdb = mvarSys.Xdb
   mvarMe.CmbIDPAI.Clear
   mvarMe.CmbIDPAI.AddItem " "
   mvarMe.CmbIDPAI.ItemData(mvarMe.CmbIDPAI.NewIndex) = 0
   Set CollCOD = New Collection
   If MyTb.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="(IDPAI=0 Or IDPAI is Null) And TPDESP=" & mvarMe.CmbTPDESP.ItemData(mvarMe.CmbTPDESP.ListIndex)) Then
      While Not MyTb.Rs.EOF
         MyTb.Popula
         mvarMe.CmbIDPAI.AddItem MyTb.DSCDESP
         mvarMe.CmbIDPAI.ItemData(mvarMe.CmbIDPAI.NewIndex) = MyTb.IDDESP
         CollCOD.Add MyTb.CODDESP, "k" & MyTb.IDDESP
         MyTb.Rs.MoveNext
      Wend
   End If
End Sub
Private Sub PopulaClasse(ByRef Queries As Collection)
   Dim oTbSLMaq As Object
   Dim bExiste As Boolean
   Dim sNewCod As String
   Dim Sql As String
   
   If TbCad Is Nothing Then Call InstanciaClasse
   If mvarMe.CmbIDPAI.ListIndex < 0 Then mvarMe.CmbIDPAI.ListIndex = 0
   With TbCad
      bExiste = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & mvarMe.TxtID.Text)
      
      sNewCod = CStr(IIf(mvarMe.CmbTPDESP.ListIndex = 0, 1, 2))
      If mvarMe.CmbIDPAI.ItemData(mvarMe.CmbIDPAI.ListIndex) <> 0 Then
         sNewCod = CollCOD("k" & mvarMe.CmbIDPAI.ItemData(mvarMe.CmbIDPAI.ListIndex))
      End If
      
      Sql = "Select Count(*) "
      Sql = Sql & " From FDESPESA "
      Sql = Sql & " Where CODDESP Like '" & sNewCod & "%'"
      Sql = Sql & " And IDLOJA=" & mvarIDLOJA
      If mvarMe.CmbIDPAI.ItemData(mvarMe.CmbIDPAI.ListIndex) = 0 Then
         Sql = Sql & " And IDPAI=0"
      Else
         Sql = Sql & " And IDDESP<>" & mvarMe.CmbIDPAI.ItemData(mvarMe.CmbIDPAI.ListIndex)
      End If
      If .IDDESP <> 0 Then
         Sql = Sql & " And IDDESP <" & .IDDESP
      End If
      If mvarSys.Xdb.AbreTabela(Sql) Then
         sNewCod = sNewCod & "." & StrZero(xVal(mvarSys.Xdb.RsAux(0) & "") + 1, 2)
      End If
      
      .IDLOJA = mvarIDLOJA
      .TPDESP = mvarMe.CmbTPDESP.ItemData(mvarMe.CmbTPDESP.ListIndex)
      .IDPAI = mvarMe.CmbIDPAI.ItemData(mvarMe.CmbIDPAI.ListIndex)
      .DSCDESP = mvarMe.TxtCampo01.Text
      .CODDESP = sNewCod
      
      If .isDirt Or Not bExiste Then
         Queries.Add .QrySave
      End If
   End With
End Sub
Private Function Salvar() As Boolean
   Dim bOk     As Boolean
   Dim xNode   As TreeViewNode
   Dim bIsDirt As Boolean
   Dim Queries As Collection
   Dim nId     As Integer
   
   If ValidaCampos() Then
      Set Queries = New Collection
      Call PopulaClasse(Queries)
      
      bOk = True
      bIsDirt = TbCad.isDirt
      If bIsDirt And Queries.Count <> 0 Then
         If mvarSys.Xdb.Executa(Queries) Then
            mvarID = TbCad.IDDESP
            mvarMe.TxtID.Text = mvarID
            If mvarMe.CmbIDPAI.ListIndex >= 0 Then
               nId = mvarMe.CmbIDPAI.ItemData(mvarMe.CmbIDPAI.ListIndex)
               Call PopulaCmbIDPAI
               Call LocalizarCombo(mvarMe.CmbIDPAI, CStr(nId), True, True)
            End If
         Else
            bOk = False
         End If
      End If
      Call ExibirResultado(mvarSys, bOk, IIf(bOk, 1, 2))
   End If
   Salvar = bOk
End Function
Private Function VerificarDependencias()
   Dim bResult    As Boolean
   Dim sMsg       As String
   
   bResult = True
   
   VerificarDependencias = bResult
Saida:
End Function

Private Sub mvarMe_CmbTPDESPClick()
   Call PopulaCmbIDPAI
End Sub

Private Sub mvarMe_CmdExcluirClick()
   If mvarID <> 0 Then
      If mvarMe.Visible Then
         Call PopulaTela
      End If
      Call Excluir
   End If
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
Private Sub mvarMe_TxtIDLostFocus()
   mvarID = xVal(mvarMe.TxtID.Text)
   Call PopulaTela
   'mvarMe.
End Sub


