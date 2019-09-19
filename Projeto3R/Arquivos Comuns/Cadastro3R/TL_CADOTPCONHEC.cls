VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADOTPCONHEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'** Palavras a serem substituidas
'** CADOTPCONHEC
'** OTPCONHEC
'** IDTPCONHEC
'** NMCONHEC

Option Explicit
Const TbName = "OTPCONHEC"
Const IdName = "IDTPCONHEC"
Const Campo1 = "NMCONHEC"
Const PergExc = "Confirma a exclusão do tipo de conhecimento?"
Const Titulo = "Tipo de Conhecimento"

Private WithEvents mvarMe  As FrmCADOTPCONHEC
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
      .Titulo = Titulo
      
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
   Set mvarMe = New FrmCADOTPCONHEC
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
   Set TbCad.xdb = mvarSys.xdb
End Sub
Private Sub PopulaTela()
   If TbCad Is Nothing Then
      Call InstanciaClasse
   End If
   
   With TbCad
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & mvarID) Then
         mvarMe.CmdExcluir.Enabled = True
         
         mvarMe.TxtID.Text = .IDTPCONHEC
         mvarMe.TxtCampo01.Text = .NMCONHEC

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
            Queries.Add TbCad.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDTPCONHEC:=xVal(mvarMe.TxtID))
            
            If mvarSys.xdb.Executa(Queries) Then
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
   If xVal(mvarMe.TxtID.Text) < 0 Then
      ExibirAviso "Código inválido."
      bResult = False
      mvarMe.TxtID.SetFocus
      GoTo Fim
   End If
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
      bExiste = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & mvarMe.TxtID.Text)
      .IDTPCONHEC = xVal(mvarMe.TxtID.Text)
      .IDLOJA = mvarIDLOJA
      .NMCONHEC = mvarMe.TxtCampo01.Text

      'If .IDTPCONHEC = 0 Then .DTCADASTRO = mvarSys.Xdb.SysDate()
      'If .isDirt Or Not bExiste Then   Queries.Add .QrySave
   End With
End Sub
Private Function Salvar() As Boolean
   Dim bOk     As Boolean
   Dim Queries As Collection
   
   If ValidaCampos() Then
      Set Queries = New Collection
      Call PopulaClasse(Queries)
      
      bOk = True
      If TbCad.isDirt Then
         bOk = TbCad.Salvar()
         If bOk Then
            mvarID = TbCad.IDTPCONHEC
            mvarMe.TxtID.Text = mvarID
         End If
      End If
      If bOk And Queries.Count > 0 Then
         bOk = mvarSys.xdb.Executa(Queries)
      End If
      Call ExibirResultado(mvarSys, bOk, IIf(bOk, 1, 2))
   End If
   Salvar = bOk
End Function
Private Function VerificarDependencias()
   Dim bResult    As Boolean
   Dim sMsg       As String
'   Dim TbAtend    As Object
'   Set TbAtend = CriarObjeto("BANCO_3R.TB_OSERVICOEVT")
'   With TbAtend
'      Set .Xdb = mvarSys.Xdb
'      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA,Ch_Where:=IdName & "=" & SqlNum(mvarMe.TxtID)) Then
'         sMsg = "Este item já possui compromisso associado e não pode ser excluído."
'         Call ExibirAviso(sMsg)
'         GoTo Saida
'      End If
'   End With
   
   bResult = True
   
   VerificarDependencias = bResult
Saida:
'   Set TbAtend = Nothing
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
   Dim SQL As String
   If xVal(mvarMe.TxtID) = 0 Or Trim(mvarMe.TxtCampo01) = "" Then
      SQL = "Update OTPCONHEC " & vbNewLine
      SQL = SQL & " Set IDTPCONHEC=0" & vbNewLine
      SQL = SQL & " , NMCONHEC = 'Não Informado'" & vbNewLine
      SQL = SQL & " Where Isnull(IDTPCONHEC,0)<= 0" & vbNewLine
      Call mvarSys.xdb.Executa(SQL)
     
      Call LimpaTela
      
      SQL = "Select Max(IDTPCONHEC)+1" & vbNewLine
      SQL = SQL & " From OTPCONHEC" & vbNewLine
      SQL = SQL & " Where IDLOJA =" & mvarIDLOJA & vbNewLine
      If mvarSys.xdb.AbreTabela(SQL) Then
         mvarMe.TxtID.Text = mvarSys.xdb.RsAux(0)
      End If
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


