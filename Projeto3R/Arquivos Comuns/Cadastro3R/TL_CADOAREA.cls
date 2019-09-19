VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADOAREA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'** CADOAREA
'** OAREA
'** IDAREA
'** DSCAREA
'** Áreas

Option Explicit
Const TbName = "OAREA"
Const IdName = "IDAREA"
Const Campo1 = "DSCAREA"

Private WithEvents mvarMe  As FrmCADOAREA
Attribute mvarMe.VB_VarHelpID = -1

Private mvarShortBar       As TL_ShortBar
Private mvarRptRec         As ReportRow
Private TbCad              As Object
Private mvarSys            As Object
Private mvarIDLOJA         As Integer
Private mvarID As Integer
Public Property Set TLCadastro(ByRef vData As Object)
   Dim SQL  As String
   Dim sAux As String
   With vData
      .Titulo = "Áreas"
      
      SQL = "Select " & IdName & " [Id.]"
      SQL = SQL & ", " & Campo1 & " [Descrição]"
      SQL = SQL & " From " & TbName
      SQL = SQL & " Where IDLOJA=" & mvarIDLOJA
      .Consulta = SQL
      
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
   Set mvarMe = New FrmCADOAREA
End Sub
Private Sub Class_Terminate()
   mvarID = 0
End Sub

Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call SetTag(mvarMe, "1VEZ", 0)
      'Call MontaTela
      'Call PopulaTela
      mvarMe.TxtCampo01.SetFocus
   End If
End Sub
Private Sub InstanciaClasse()
   Set TbCad = Nothing
   Set TbCad = CriarObjeto("BANCO_3R.TB_" & TbName)
   Set TbCad.xdb = mvarSys.xdb
End Sub
Private Sub PopulaTela()
   Dim bExiste As Boolean
   If TbCad Is Nothing Then
      Call InstanciaClasse
   End If
   
   With TbCad
      bExiste = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & mvarID)
      If mvarMe.TxtID.Text = "" And mvarID = 0 Then bExiste = False
      
      If bExiste Then
         mvarMe.CmdExcluir.Enabled = True
         
         mvarMe.TxtID.Text = .IDAREA
         mvarMe.TxtCampo01.Text = .DSCAREA
         mvarMe.ChkATIVO.Value = .ATIVO
      Else
         Call LimpaTela
      End If
   End With
End Sub
Public Function Excluir() As Boolean
   Dim Queries As Collection
   
   
   If mvarMe.TxtID <> "" Then
      If VerificarDependencias() Then
         If ExibirPergunta("Confirma a exclusão do área?", "Atenção", False) = vbYes Then
            Call TbCad.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & xVal(mvarMe.TxtID))
            
            Set Queries = New Collection
            Queries.Add TbCad.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDAREA:=xVal(mvarMe.TxtID))
            
            If mvarSys.xdb.Executa(Queries) Then
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
         If n Is mvarMe.TxtID Then
            If mvarMe.TxtID.Text = "0" Then
               mvarMe.TxtID.Tag = "0"
            Else
               mvarMe.TxtID.Tag = ""
               n.Text = ""
            End If
         Else
            n.Text = ""
         End If
      Next
   End With
End Sub
Private Sub MontaTela()
   Dim MyTb As Object
   
End Sub
Private Sub PopulaClasse(ByRef Queries As Collection)
   Dim oTbSLMaq As Object
   Dim bExiste As Boolean
   Dim SQL As String
   
   'If TbCad Is Nothing Then
   Call InstanciaClasse
   
   With TbCad
      bExiste = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & mvarMe.TxtID.Text)
      If .IDAREA = 0 Then
'         .DTCADASTRO = mvarSys.Xdb.SysDate()
      End If
      .IDLOJA = mvarIDLOJA
      .DSCAREA = mvarMe.TxtCampo01.Text
      .ATIVO = mvarMe.ChkATIVO.Value
      If .isDirt And Not bExiste And mvarMe.TxtID.Text = "0" Then
         SQL = "Insert Into OAREA (IDLOJA, IDAREA, DSCAREA, ALTERSTAMP, TIMESTAMP, ATIVO)"
         SQL = SQL & " Values ("
         SQL = SQL & SqlNum(.IDLOJA)
         SQL = SQL & ", " & SqlNum(.IDAREA)
         SQL = SQL & ", " & SqlStr(.DSCAREA)
         SQL = SQL & ", 1"
         SQL = SQL & ", GetDate()"
         SQL = SQL & ", " & SqlNum(.ATIVO)
         SQL = SQL & ")"
'         Set Queries = New Collection
         Queries.Add SQL
      End If
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
         If Queries.Count > 0 Then
            Call mvarSys.xdb.Executa(Queries)
         End If
         If TbCad.Salvar Then
            bOk = True
            mvarID = TbCad.IDAREA
            mvarMe.TxtID.Text = mvarID
         End If
      Else
         bOk = True
      End If
      
      'bOk = False
      'If Queries.Count = 0 Then
      '   bOk = True
      'Else
      '   If mvarSys.xDb.Executa(Queries) Then
      '      bOk = True
      '   End If
      'End If
      
      Call ExibirResultado(mvarSys, bOk, IIf(bOk, 1, 2))
   End If
   Salvar = bOk
End Function
Private Function VerificarDependencias()
   Dim bResult    As Boolean
   Dim TbServ    As Object
   Dim sMsg       As String
   
   Set TbServ = CriarObjeto("BANCO_3R.TB_OSERVICOEVT")
   With TbServ
      Set .xdb = mvarSys.xdb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & SqlNum(mvarMe.TxtID)) Then
         sMsg = "Este item já possui compromisso associado e não pode ser excluído."
         Call ExibirAviso(sMsg)
         GoTo Saida
      End If
   End With
   
   bResult = True
   
   VerificarDependencias = bResult
Saida:
   Set TbServ = Nothing
End Function
Private Sub mvarMe_CmdExcluirClick()
   'If mvarID <> 0 Then
   If mvarMe.TxtID.Text <> "" Then
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
   mvarID = 0
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


