VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADOTPSERVICO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe  As FrmCADOTPSERVICO
Attribute mvarMe.VB_VarHelpID = -1

Private mvarShortBar       As TL_ShortBar
Private mvarRptRec         As ReportRow
Private TbCad              As Object
Private mvarSys            As Object
Private mvarIDLOJA         As Integer

Private mvarIDTPSERVICO    As Integer
Public Property Set TLCadastro(ByRef vData As Object)
   Dim sAux As String
   Dim Sql  As String
   With vData
      .Titulo = "Serviços"
      
      Sql = "Select IDTPSERVICO [Id.], DSCSERVICO [Descrição] "
      Sql = Sql & " From OTPSERVICO"
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      .Consulta = Sql
      
      sAux = ""
      sAux = sAux & "|CAMPO=IDTPSERVICO|CHAVE=S|GRUPO=0|VISIBLE=S:"
      .TAGCAMPOS = sAux
   End With
End Property
Public Property Let IDTPSERVICO(ByVal vData As String)
   mvarIDTPSERVICO = vData
End Property
Public Property Get IDTPSERVICO() As String
   IDTPSERVICO = mvarIDTPSERVICO
End Property
Public Property Set RptRec(ByVal vData As ReportRow)
   Set mvarRptRec = vData
'   mvarIDLOJA = RptRec.Record.Item(0).Value
   mvarIDTPSERVICO = RptRec.Record.Item(0).Value
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
Public Property Let IDLOJA(ByVal vData As String)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As String
   IDLOJA = mvarIDLOJA
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmCADOTPSERVICO
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
   Set TbCad = CriarObjeto("BANCO_3R.TB_OTPSERVICO")
   Set TbCad.Xdb = mvarSys.Xdb
End Sub
Private Sub PopulaTela()
   If TbCad Is Nothing Then
      Call InstanciaClasse
   End If
   
   With TbCad
      If .Pesquisar(Ch_IDTPSERVICO:=mvarIDTPSERVICO) Then
         mvarMe.CmdExcluir.Enabled = True
         
         mvarMe.TxtID.Text = .IDTPSERVICO
         mvarMe.TxtCampo01.Text = .DSCSERVICO
         mvarMe.ChkATIVO.Value = .ATIVO
         
      Else
         Call LimpaTela
      End If
   End With
End Sub
Public Function Excluir() As Boolean
   Dim Queries As Collection
   
   
   If mvarIDTPSERVICO <> 0 Then
      If VerificarDependencias() Then
         If ExibirPergunta("Confirma a exclusão do serviço?", "Atenção", False) = vbYes Then
            Call TbCad.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDTPSERVICO:=xVal(mvarMe.TxtID))
            
            Set Queries = New Collection
            'Queries.Add "Delete From OSALA_MAQUINA Where IDMAQUINA=" & SqlNum(mvarMe.TxtID)
            Queries.Add TbCad.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDTPSERVICO:=xVal(mvarMe.TxtID))
            
            If mvarSys.Xdb.Executa(Queries) Then
               Call ExibirResultado(mvarSys, True)
               
               'mvarIsDirt = True
               
               Call LimpaTela
               If Not mvarShortBar Is Nothing Then
                  mvarShortBar.TLCadastro.F_Refresh
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
         n.Text = ""
      Next
   End With
End Sub
Private Sub MontaTela()
   Dim MyTb As Object
   
End Sub
Private Sub PopulaClasse(ByRef Queries As Collection)
   Dim oTbSLMaq As Object
   Dim bExiste As Boolean
   
   If TbCad Is Nothing Then Call InstanciaClasse
   
   With TbCad
      bExiste = .Pesquisar(Ch_IDTPSERVICO:=mvarMe.TxtID.Text)
      If .IDTPSERVICO = 0 Then
'         .DTCADASTRO = mvarSys.Xdb.SysDate()
      End If
      .IDLOJA = mvarIDLOJA
      .DSCSERVICO = mvarMe.TxtCampo01.Text
      .ATIVO = mvarMe.ChkATIVO.Value
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
      
      bOk = True
      bIsDirt = TbCad.isDirt
      If bIsDirt Then
         bOk = TbCad.Salvar()
         If bOk Then
            mvarIDTPSERVICO = TbCad.IDTPSERVICO
            mvarMe.TxtID.Text = mvarIDTPSERVICO
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
   Dim TbAtend    As Object
   Dim sMsg       As String
   
   Set TbAtend = CriarObjeto("BANCO_3R.TB_OSERVICOEVT")
   With TbAtend
      Set .Xdb = mvarSys.Xdb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDTPSERVICO=" & SqlNum(mvarMe.TxtID)) Then
         sMsg = "Este item já possui compromisso associado e não pode ser excluído."
         Call ExibirAviso(sMsg)
         GoTo Saida
      End If
   End With
   
   bResult = True
   
   VerificarDependencias = bResult
Saida:
   Set TbAtend = Nothing
End Function
Private Sub mvarMe_CmdExcluirClick()
   If mvarIDTPSERVICO <> 0 Then
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
   mvarIDTPSERVICO = xVal(mvarMe.TxtID.Text)
   Call PopulaTela
   'mvarMe.
End Sub
