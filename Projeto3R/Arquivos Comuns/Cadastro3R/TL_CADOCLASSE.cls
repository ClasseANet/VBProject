VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADOCLASSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'** Palavras a serem substituidas
'** CADOCLASSE
'** OCLASSE
'** IDCLASSE
'** DSCCLASSE

Option Explicit
Const TbName = "OCLASSE"
Const IdName = "IDCLASSE"
Const Campo1 = "DSCCLASSE"
Const PergExc = "Confirma a exclusão da classe?"
Const Titulo = "Classes"

Private WithEvents mvarMe  As FrmCADOCLASSE
Attribute mvarMe.VB_VarHelpID = -1

Private mvarShortBar       As TL_ShortBar
Private mvarRptRec         As ReportRow
Private TbCad              As Object
Private mvarSys            As Object
Private mvarIDLOJA         As Integer
Private mvarID As Integer
Public Property Set TLCadastro(ByRef vData As Object)
   Dim Sql  As String
   Dim sAux As String
   With vData
      .Titulo = Titulo
      
      Sql = "Select " & IdName & " [Id.]"
      Sql = Sql & ", " & Campo1 & " [Descrição]"
      Sql = Sql & " From " & TbName
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      .Consulta = Sql
      
      sAux = ""
      sAux = sAux & "|CAMPO=" & IdName & "|CHAVE=S|GRUPO=0|VISIBLE=S:"
      .TAGCAMPOS = sAux
   End With
End Property
Public Property Let ID(ByVal vData As String)
   mvarID = vData
End Property
Public Property Get ID() As String
   ID = mvarID
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
   Set mvarMe = New FrmCADOCLASSE
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
      If .Pesquisar(Ch_Where:=IdName & "=" & mvarID & " And IDLOJA=" & mvarIDLOJA) Then
         mvarMe.CmdExcluir.Enabled = True
         
         mvarMe.TxtID.Text = .IDCLASSE
         mvarMe.TxtCampo01.Text = .DSCCLASSE
         mvarMe.ChkATIVO.Value = .ATIVO
         Call LocalizarCombo(mvarMe.CmbIDPAI, .IDPAI, True, True)

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
            Queries.Add TbCad.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCLASSE:=xVal(mvarMe.TxtID))
            
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
   Call PopulaCmbIDPAI
End Sub
Private Sub PopulaCmbIDPAI()
   Dim MyTb As Object
   Set MyTb = CriarObjeto("BANCO_3R.TB_OCLASSE")
   Set MyTb.Xdb = mvarSys.Xdb
   mvarMe.CmbIDPAI.Clear
   mvarMe.CmbIDPAI.AddItem " "
   mvarMe.CmbIDPAI.ItemData(mvarMe.CmbIDPAI.NewIndex) = 0
   If MyTb.Pesquisar(Ch_Where:="(IDPAI=0 Or IDPAI is Null) And IDLOJA=" & mvarIDLOJA) Then
      While Not MyTb.Rs.EOF
         MyTb.Popula
         mvarMe.CmbIDPAI.AddItem MyTb.DSCCLASSE
         mvarMe.CmbIDPAI.ItemData(mvarMe.CmbIDPAI.NewIndex) = MyTb.IDCLASSE
         MyTb.Rs.MoveNext
      Wend
   End If
End Sub
Private Sub PopulaClasse(ByRef Queries As Collection)
   Dim oTbSLMaq As Object
   Dim bExiste As Boolean
   
   If TbCad Is Nothing Then Call InstanciaClasse
   If mvarMe.CmbIDPAI.ListIndex < 0 Then mvarMe.CmbIDPAI.ListIndex = 0
   
   With TbCad
      bExiste = .Pesquisar(Ch_Where:=IdName & "=" & mvarMe.TxtID.Text & " And IDLOJA=" & mvarIDLOJA)
      .IDLOJA = mvarIDLOJA
      .DSCCLASSE = mvarMe.TxtCampo01.Text
      .ATIVO = mvarMe.ChkATIVO.Value
      .IDPAI = mvarMe.CmbIDPAI.ItemData(mvarMe.CmbIDPAI.ListIndex)
      'If .IDTPTRATAMENTO = 0 Then .DTCADASTRO = mvarSys.Xdb.SysDate()
      'If .isDirt Or Not bExiste Then   Queries.Add .QrySave
   End With
End Sub
Private Function Salvar() As Boolean
   Dim bOk     As Boolean
   Dim Queries As Collection
   Dim nId     As Integer
   
   If ValidaCampos() Then
      Set Queries = New Collection
      Call PopulaClasse(Queries)
     
      bOk = True
      If TbCad.isDirt Then
         bOk = TbCad.Salvar()
         If bOk Then
            mvarID = TbCad.IDCLASSE
            mvarMe.TxtID.Text = mvarID
            
            If mvarMe.CmbIDPAI.ListIndex <= 0 Then
               nId = mvarMe.CmbIDPAI.ItemData(mvarMe.CmbIDPAI.ListIndex)
               Call PopulaCmbIDPAI
               Call LocalizarCombo(mvarMe.CmbIDPAI, CStr(nId), True, True)
            End If
            
         End If
      End If
      If bOk And Queries.Count > 0 Then
         bOk = mvarSys.Xdb.Executa(Queries)
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


