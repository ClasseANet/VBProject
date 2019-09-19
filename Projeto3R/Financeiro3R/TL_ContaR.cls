VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_ContaR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe    As FrmContaR
Attribute mvarMe.VB_VarHelpID = -1

Dim nLenIDVenda   As Integer
Dim dToday        As Date

Private mvarSys            As Object
Private mvarIDLOJA         As Integer
Private mvarIDATENDIMENTO  As Long
Private mvarIDCLIENTE      As Long
Private mvarIDCONTA        As Long
Private mvarCalControl     As CalendarControl
Private mvarCalEvent       As CalendarEvent
Public Property Set CalEvent(ByRef vData As Object)
   Set mvarCalEvent = vData
End Property
Public Property Get CalEvent() As Object
    Set CalEvent = mvarCalEvent
End Property
Public Property Set CalControl(ByVal vData As Object)
   Set mvarCalControl = vData
   On Error Resume Next
   mvarCalControl.DataProvider.ClearCache
End Property
Public Property Get CalControl() As Object
    Set CalControl = mvarCalControl
End Property
Public Property Let IDCONTA(ByVal vData As Long)
    mvarIDCONTA = vData
End Property
Public Property Get IDCONTA() As Long
   IDCONTA = mvarIDCONTA
End Property
Public Property Let IDCLIENTE(ByVal vData As Long)
    mvarIDCLIENTE = vData
End Property
Public Property Get IDCLIENTE() As Long
   IDCLIENTE = mvarIDCLIENTE
End Property
Public Property Let IDATENDIMENTO(ByVal vData As Long)
    mvarIDATENDIMENTO = vData
End Property
Public Property Get IDATENDIMENTO() As Long
   IDATENDIMENTO = mvarIDATENDIMENTO
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
    mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   
   If mvarIDLOJA = 0 Then mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
   nLenIDVenda = mvarSys.Propriedades("LENIDVENDA")
   On Error Resume Next
   dToday = CDate(Format(mvarSys.xDb.sysdate, "DD/MM/YYYY"))
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmContaR
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
End Sub
Public Sub Show(Optional pMode As FormShowConstants = vbModal)
'   If mvarIDLOJA = 0 Then
'      Call ExibirAviso("Código da Loja Inválido! (""" & mvarIDLOJA & """)")
'      Exit Sub
'   End If
   
   mvarMe.Show pMode
End Sub
Private Sub mvarMe_Activate()
   'Call PopulaTela
   
   If GetTag(mvarMe, "1VEZ", "1") = "1" Then
      Call SetTag(mvarMe, "1VEZ", "0")
   End If
End Sub
Private Sub mvarMe_CmdCancelClick()
   Unload mvarMe
End Sub

Private Sub mvarMe_CmdExcluirClick()
   Dim sMsg As String
   Dim Queries As Collection
   Dim Sql As String
   Dim nRecibo As Long
   Dim MyRs As Object
   
   Dim TBContaR As Object 'TB_FCONTAR
   
   sMsg = "Esta operação irá excluir definitivamente a conta '" & mvarMe.TxtIDCONTA.Text & "'"
   sMsg = sMsg & vbNewLine & vbNewLine
   sMsg = sMsg & "Deseja continuar?"
   If ExibirPergunta(sMsg, , False) = vbYes Then
      If InputBoxPassword("Informe a senha.") = mvarSys.Propriedades("SENHAMESTRE") Then
         Set Queries = New Collection
         
         Set TBContaR = CriarObjeto("BANCO_3R.TB_FCONTAR")
         With TBContaR
            Set .xDb = mvarSys.xDb
            If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCONTA:=mvarIDCONTA) Then
               Queries.Add .QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCONTA:=mvarIDCONTA)
               If mvarSys.xDb.Executa(Queries) Then
                  Call ExibirResultado(mvarSys, True)
                  Unload mvarMe
               Else
                  Call ExibirResultado(mvarSys, False)
               End If
            End If
         End With
      Else
         Call ExibirAviso("Esta senha não permite a operação.")
      End If
   End If
End Sub

Private Sub mvarMe_CmdIDCLIENTEClick()
   Dim MyContato As Object
   
   Set MyContato = CriarObjeto("Contato3R.TL_CADContato")
   With MyContato
      Set .Sys = mvarSys
      .IDCLIENTE = mvarIDCLIENTE
      Call .Show
   End With
   Set MyContato = Nothing
End Sub
Private Sub mvarMe_CmdLovCliClick()
   Dim MyContato As Object
   
   Set MyContato = CriarObjeto("Contato3R.TL_NovoLov")
   With MyContato
      Set .Sys = mvarSys
      .Filtro = Trim(mvarMe.TxtNOME.Text)
      Call .ShowClientes
      If .IdSelecionado <> 0 Then
         Call PopulaCliente(.IdSelecionado)
      End If
   End With
   Set MyContato = Nothing
End Sub

Private Sub mvarMe_CmdOkClick()
   DoEvents
   If Not VerificaContaR() Then Exit Sub
   If SalvarContaR Then
      Unload mvarMe
   End If
End Sub
Private Function VerificaContaR() As Boolean
   Dim i As Long
   Dim bDin As Boolean
   'If Trim(mvarMe.TxtTEL1.Text) = "" Then
   '   Call ExibirInformacao("Informe o telefone do cliente.")
   '   mvarMe.TxtTEL1.SetFocus
   '   Exit Function
   'End If
   If Trim(mvarMe.TxtNOME.Text) = "" Then
      Call ExibirInformacao("Informe o nome do cliente.")
      mvarMe.TxtNOME.SetFocus
      Exit Function
   End If
   If GetTag(mvarMe.TxtNOME, "IDCLIENTE", "0") = "0" Then
      Call ExibirInformacao("Cliente sem identificação." & vbNewLine & "Favor verificar cadastro.")
      Exit Function
   End If
   
   If CDate(Format(mvarMe.TxtDTPREV.Value, "DD/MM/YYYY")) > dToday Then
      Call ExibirInformacao("Data da conta não pode ser maior que hoje." & vbNewLine & "Favor verificar data.")
      If mvarMe.TxtDTPREV.Enabled Then
         mvarMe.TxtDTPREV.SetFocus
      End If
      Exit Function
   End If
    
   VerificaContaR = True
End Function
Private Function SalvarContaR() As Boolean
   Dim bResult    As Boolean
   Dim TBContaR    As Object 'TB_FCONTAR
   '* Cliente
   Call SalvarCliente
   
   '* Atendimento
   Call PopulaClasse_ContaPR(TBContaR)
   bResult = TBContaR.Salvar
   
   If bResult Then
      Call ExibirInformacao("Erro ao gravar conta.")
   End If
        
Saida:
   Set TBContaR = Nothing

End Function
Private Sub SalvarCliente()
   Dim TBCliente As Object 'TB_OCLIENTE
   Dim bClienteOK As Boolean
   
   If Not (mvarMe.TxtNOME.Enabled) Then Exit Sub
   
   Set TBCliente = CriarObjeto("BANCO_3R.TB_OCLIENTE")
   With TBCliente
      Set .xDb = mvarSys.xDb
      If .Pesquisar(Ch_IDCLIENTE:=mvarIDCLIENTE, Ch_IDLOJA:=mvarIDLOJA) Then
         .IDLOJA0 = IIf(mvarSys.GetParam("OCLIENTEUNICO", Default:=0) = "1", 0, mvarIDLOJA)
         .NOME = Trim(mvarMe.TxtNOME.Text)
         If .IDCLIENTE = 0 Then .DTCADASTRO = mvarSys.xDb.sysdate
         bClienteOK = (.NOME <> "") And (.TEL1 & .TEL2 <> "")
         If .isDirt And bClienteOK Then
            Call .Alterar
            mvarIDCLIENTE = .IDCLIENTE
         End If
      End If
   End With
   
   Set TBCliente = Nothing
   
End Sub
Private Sub PopulaClasse_ContaPR(ByRef pTBContaR As Object)
   Dim sDate   As String
   
   Set pTBContaR = CriarObjeto("BANCO_3R.TB_FCONTAR")
   
   With pTBContaR
      .xDb = mvarSys.xDb
      
      Call .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCONTA:=mvarIDCONTA)
      .IDCONTA = mvarIDCONTA
      .IDLOJA = mvarIDLOJA
      
      .DTPREV = mvarMe.CmbDTPREV.Value
      .IDCLIENTE = mvarIDCLIENTE
      If .SITCONTA = "" Then
         .SITCONTA = "00"
      End If
      .Valor = xVal(mvarMe.TxtValor.Text)
      .OBS = mvarMe.TxtHISTORICO.Text
   End With
End Sub
Private Sub mvarMe_Load()
   mvarMe.Icon = mvarSys.MDI.Icon
   '* Monta Tela
   Call MontaTela
   Call PopulaTela
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub MontaTela()
   With mvarMe
      .CmdExcluir.Enabled = False
     
      .TxtIDCONTA.BackColor = &HCCEEDD
   End With
End Sub
Private Sub PopulaTela()
   Dim Sql     As String
   Dim TBAtend As Object 'TB_OATENDIMENTO
   Dim TBContaR As Object 'TB_FCONTAR
   Dim TbAV    As Object 'TB_OATENDIMENTO_VENDA
   Dim bOk     As Boolean
   
   Call LimparTela
   
   If mvarIDCONTA = 0 And mvarIDATENDIMENTO <> 0 Then
      Set TbAV = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_VENDA")
      Set TbAV.xDb = mvarSys.xDb
      
      bOk = TbAV.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO)
      If bOk Then mvarIDCONTA = TbAV.IDVENDA
   End If
   
   If mvarIDCONTA > 0 Then
      mvarMe.CmbDTPREV.Enabled = False
      If TBContaR Is Nothing Then
         Set TBContaR = CriarObjeto("BANCO_3R.TB_FCONTAR")
         Set TBContaR.xDb = mvarSys.xDb
         bOk = TBContaR.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCONTA:=mvarIDCONTA)
      End If
      If bOk Then
         With TBContaR
                     
            mvarMe.TxtIDCONTA.Text = StrZero(.IDVENDA, mvarSys.Propriedades("LENIDVENDA"))
            mvarMe.CmbDTPREV.Value = .DTVENDA
            mvarMe.TxtDTPREV.Text = .DTVENDA
                                    
            Call PopulaCliente(.IDCLIENTE)

            mvarMe.TxtValor.Text = ValBr(.VLVENDA)
            mvarMe.TxtHISTORICO.Text = .HISTORICO
         End With
      End If
      Set TBContaR = Nothing
   Else
   
   End If
End Sub
Private Sub LimparTela()
   With mvarMe
      .TxtIDCONTA.Text = ""
      .TxtIDCONTA.Enabled = False
      .CmbDTPREV.Value = mvarSys.xDb.sysdate()
      .TxtDTPREV.Text = .CmbDTPREV.Value
      
      Call SetTag(.TxtNOME, "IDCLIENTE", "")
      .TxtNOME.Text = ""
      
      .TxtValor.Text = ValBr(0)
      .TxtHISTORICO.Text = ""
   End With
End Sub
Private Sub PopulaCliente(pIDCLIENTE As Long)
   Dim TbCli   As Object 'TB_OCLIENTE
   Dim bAchou  As Boolean
   
   Set TbCli = CriarObjeto("BANCO_3R.TB_OCLIENTE")
   With TbCli
      Set .xDb = mvarSys.xDb
      bAchou = .Pesquisar(Ch_IDCLIENTE:=pIDCLIENTE)
      If bAchou Then
         mvarIDCLIENTE = .IDCLIENTE
         Call SetTag(mvarMe.TxtNOME, "IDCLIENTE", .IDCLIENTE)
         mvarMe.TxtNOME.Text = .NOME
      Else
         mvarIDCLIENTE = 0
         Call SetTag(mvarMe.TxtNOME, "IDCLIENTE", "")
         mvarMe.TxtNOME.Text = ""
      End If
   End With
   Set TbCli = Nothing
End Sub
Private Sub mvarMe_TxtIDCONTALostFocus()
   If Val(mvarMe.TxtIDCONTA.Text) = 0 Then
      mvarMe.TxtIDCONTA.Text = ""
   Else
      mvarMe.TxtIDCONTA.Text = StrZero(mvarMe.TxtIDCONTA.Text, mvarSys.Propriedades("LENIDVENDA"))
   End If
End Sub




