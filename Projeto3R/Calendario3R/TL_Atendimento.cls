VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Atendimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe    As FrmAtendimento
Attribute mvarMe.VB_VarHelpID = -1
Private WithEvents ContextMenu   As CommandBars
Attribute ContextMenu.VB_VarHelpID = -1

Private bAtendFechado As Boolean
Private bExisteVenda  As Boolean
Private bVendaFechada As Boolean
Private bChave        As Boolean
Private bFLGDISPARO   As Boolean
Private mvarTbAtend    As Object
Private bIsFOTO       As Boolean
Private nIDTRAT       As Integer

Private mvarCalControl     As CalendarControl
Private mvarSys            As Object
Private mvarIDLOJA         As Integer
Private mvarIDSALA         As Integer
Private mvarIDMAQUINA      As Integer
Private mvarIDAGENDA       As Integer
Private mvarIDEVENTO       As Long
Private mvarIDATENDIMENTO  As Long
Private mvarIDCLIENTE      As Long
Private mvarCalEvent       As CalendarEvent
Public Property Set CalEvent(ByRef vData As Object)
   Set mvarCalEvent = vData
End Property
Public Property Get CalEvent() As Object
    Set CalEvent = mvarCalEvent
End Property
Public Property Set CalControl(ByVal vData As Object)
   Set mvarCalControl = vData
   If Not mvarCalControl Is Nothing Then
      mvarCalControl.DataProvider.ClearCache
   End If
End Property
Public Property Get CalControl() As Object
    Set CalControl = mvarCalControl
End Property
Public Property Let IDCLIENTE(ByVal vData As Long)
    mvarIDCLIENTE = vData
End Property
Public Property Get IDCLIENTE() As Long
   IDCLIENTE = mvarIDCLIENTE
End Property
Public Property Let IDATENDIMENTO(ByVal vData As Long)
   If mvarIDEVENTO = 0 Then
      If mvarTbAtend Is Nothing Then
         Set mvarTbAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
      End If
      Set mvarTbAtend.xDb = mvarSys.xDb
      If mvarTbAtend.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=vData) Then
         mvarIDEVENTO = mvarTbAtend.IDEVENTO
      End If
   End If
   mvarIDATENDIMENTO = vData
End Property
Public Property Get IDATENDIMENTO() As Long
   IDATENDIMENTO = mvarIDATENDIMENTO
End Property
Public Property Let IDEVENTO(ByVal vData As Long)
   If mvarIDATENDIMENTO = 0 Then
      If mvarTbAtend Is Nothing Then
         Set mvarTbAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
      End If
      Set mvarTbAtend.xDb = mvarSys.xDb
      If mvarTbAtend.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDEVENTO=" & vData) Then
         mvarIDATENDIMENTO = mvarTbAtend.IDATENDIMENTO
         mvarIDSALA = mvarTbAtend.IDSALA
         mvarIDMAQUINA = mvarTbAtend.IDMAQUINA
         mvarIDCLIENTE = mvarTbAtend.IDCLIENTE
      End If
   End If
    
   mvarIDEVENTO = vData
End Property
Public Property Get IDEVENTO() As Long
   IDEVENTO = mvarIDEVENTO
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
    mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDSALA(ByVal vData As Integer)
    mvarIDSALA = vData
    
   Dim Sql As String
   Sql = "Select M.IDMAQUINA "
   Sql = Sql & " From OMAQUINA M "
   Sql = Sql & " Join OSALA_MAQUINA SM On M.IDLOJA=SM.IDLOJA And M.IDMAQUINA=SM.IDMAQUINA"
   Sql = Sql & " Join OSALA S On SM.IDLOJA=S.IDLOJA And SM.IDSALA=S.IDSALA"
   Sql = Sql & " Where M.SITMAQUINA=1"
   Sql = Sql & " And S.IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And S.IDSALA=" & mvarIDSALA
   Sql = Sql & " And M.SITMAQUINA=1"
   If mvarSys.xDb.AbreTabela(Sql) Then
      mvarIDMAQUINA = mvarSys.xDb.RsAux("IDMAQUINA")
   End If
       
End Property
Public Property Get IDSALA() As Integer
   IDSALA = mvarIDSALA
End Property
Public Property Let IDMAQUINA(ByVal vData As Integer)
    mvarIDMAQUINA = vData
End Property
Public Property Get IDMAQUINA() As Integer
   IDMAQUINA = mvarIDMAQUINA
End Property

Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   If mvarIDLOJA = 0 Then mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmAtendimento
   mvarIDCLIENTE = 0
   mvarIDAGENDA = 1
   bChave = False
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
   Set mvarCalControl = Nothing
End Sub
Public Sub Show(Optional pMode As FormShowConstants = vbModal)
   If mvarIDLOJA = 0 Then
      Call ExibirAviso("Código da Loja Inválido! (""" & mvarIDLOJA & """)")
      Exit Sub
   End If
   mvarMe.Show pMode
End Sub
Public Sub Limpar()
   mvarIDEVENTO = 0
   mvarIDATENDIMENTO = 0
   mvarIDCLIENTE = 0
End Sub
Private Sub ContextMenu_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   If Control.Category = "ATENDIMENTO" Then
      mvarMe.MousePointer = vbHourglass
      mvarMe.CmdVenda.Enabled = False
   
      Select Case Control.Id
         Case 1: Call mvarMe_CmdVendaClick
         Case 2: Call ExibirFatura
      End Select
      
      mvarMe.MousePointer = vbDefault
      mvarMe.CmdVenda.Enabled = True
   End If
End Sub

Private Sub mvarMe_Activate()
   mvarMe.GrdSESSAO.Draw
'mvarMe_Load
'* Monta Tela
'Call MontaTela
'* Popula Tela
'Call CarregaFechamento
'   Call PopulaTela

   If GetTag(mvarMe, "1VEZ", "1") = "1" Then
      Call SetTag(mvarMe, "1VEZ", "0")
      If mvarMe.TxtTEL1.Text = "" And mvarMe.TxtTEL1.Enabled Then
         mvarMe.TxtTEL1.SetFocus
      Else
         If mvarMe.CmbIDFUNCIONARIO.ListIndex = -1 Then
            If mvarMe.CmbIDFUNCIONARIO.Enabled And mvarMe.CmbIDFUNCIONARIO.Visible Then
               mvarMe.CmbIDFUNCIONARIO.SetFocus
            End If
         Else
            mvarMe.GrdSESSAO.SetFocus
         End If
      End If
      
      Call VerificaMensagemI
      Screen.MousePointer = vbDefault
   End If
End Sub
Private Sub VerificaMensagemI()
   Dim Sql As String
   Dim sMsg As String
   
   If mvarIDATENDIMENTO <= 0 Then Exit Sub
   
   Sql = "Select IDFATURA "
   Sql = Sql & " From FFATURA"
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And IDATENDIMENTO=" & SqlNum(mvarIDATENDIMENTO)
   Sql = Sql & " And SITFATURA=0"
   
   With mvarSys.xDb
      If .AbreTabela(Sql) Then
         With .RsAux
            If .RecordCount > 0 Then
               sMsg = "Doc.: " & StrZero(mvarIDATENDIMENTO, 6)
               sMsg = Mid(sMsg, 1, 28) & IIf(Len(sMsg) > 30, "...", "") & vbNewLine
               sMsg = sMsg & "Este Atendimento possui fatura 'Em Aberto'"
               Call ExibirMensagemI(mvarSys, sMsg)
            End If
         End With
      End If
   End With
End Sub
Private Sub mvarMe_CmbHHFIMChange()
'   DoEvents
   With mvarMe
      If IsNull(.CmbHHFIM.Value) Then
         If mvarCalEvent Is Nothing Then
            .CmbHHFIM.Value = .CmbHHINI.Value
         Else
            .CmbHHFIM.Value = mvarCalEvent.EndTime
         End If
         .TxtHHFIM.Text = Format(.CmbHHFIM.Value, "hh:mm")
         .CmbHHFIM.Value = Null
         Call SetTag(.CmbHHFIM, "NULL", 1)
      Else
         '.CmbHHFIM.Value = Format(.CmbDTATEND.Value, "dd/mm/yyyy") & " " & Format(.CmbHHFIM.Value, "hh:mm")
         If GetTag(.CmbHHFIM, "NULL", 0) = 1 Then
            Call SetTag(.CmbHHFIM, "NULL", 0)
            .CmbHHFIM.Value = Format(.CmbDTATEND.Value, "dd/mm/yyyy") & " " & mvarSys.xDb.Sysdate(2)
            .TxtHHFIM.Text = Format(.CmbHHFIM.Value, "hh:mm")
         End If
      End If
      .TxtHHFIM.Enabled = Not IsNull(.CmbHHFIM.Value)
   End With
   Call ExibeDuracao
End Sub
Private Sub mvarMe_CmbHHFIMValidate(Cancel As Boolean)
   With mvarMe
      If mvarMe.ActiveControl Is mvarMe.CmdCancel Then
         Exit Sub
      End If
      If Not IsNull(.CmbHHFIM.Value) And Not IsNull(.CmbHHINI.Value) Then
         Cancel = (Format(.CmbHHFIM.Value, "hh:mm") <= Format(.CmbHHINI.Value, "hh:mm"))
         If Cancel Then
            MsgBox "Hora de Término Inválida.", vbOKOnly + vbInformation, "Hora"
            .CmbHHFIM.Value = .CmbHHINI.Value
            '.CmbHHFIM.SetFocus
            .TxtHHFIM.Text = .CmbHHFIM.Value
            .TxtHHFIM.Enabled = True
            .TxtHHFIM.SetFocus
         End If
      End If
   End With
End Sub

Private Sub mvarMe_CmbHHINIChange()
   With mvarMe
      
      If IsNull(.CmbHHINI.Value) Then
         .CmbHHINI.Value = Format(.CmbDTATEND.Value, "dd/mm/yyyy") & " " & mvarSys.xDb.Sysdate(2)
         .TxtHHINI.Text = Format(.CmbHHINI.Value, "hh:mm")
         .CmbHHINI.Value = Null
         
         .TxtHHFIM.Text = Format(DateAdd("n", DateDiff("n", mvarCalEvent.StartTime, mvarCalEvent.EndTime), CDate(.TxtHHINI.Text)), "hh:mm")
         .CmbHHFIM.Value = Format(.TxtHHFIM.Text, "hh:mm")
         
         
         
      Else
         If Format(.CmbHHINI.Value, "dd/mm/yyyy") <> Format(.CmbDTATEND.Value, "dd/mm/yyyy") Then
            .CmbHHINI.Value = Format(.CmbDTATEND.Value, "dd/mm/yyyy") & " " & Format(.CmbHHINI.Value, "hh:mm")
            .TxtHHINI.Text = Format(.CmbHHINI.Value, "hh:mm")
         End If
      End If
      If IsNull(.CmbHHFIM.Value) Then
         .CmbHHFIM.Value = .CmbHHINI.Value
         .TxtHHFIM.Text = Format(.CmbHHFIM.Value, "hh:mm")
         .CmbHHFIM.Value = Null
      Else
         If Format(.CmbHHFIM.Value, "dd/mm/yyyy") <> Format(.CmbDTATEND.Value, "dd/mm/yyyy") Then
            .CmbHHFIM.Value = Format(.CmbDTATEND.Value, "dd/mm/yyyy") & " " & Format(.CmbHHFIM.Value, "hh:mm")
            .TxtHHFIM.Text = Format(.CmbHHFIM.Value, "hh:mm")
         End If
         While .CmbHHFIM.Value <= .CmbHHINI.Value
            .CmbHHFIM.Value = DateAdd("h", 1, .CmbHHFIM.Value)
            .TxtHHFIM.Text = Format(.CmbHHFIM.Value, "hh:mm")
         Wend
      End If
      .TxtHHINI.Enabled = Not IsNull(.CmbHHINI)
   End With
   Call ExibeDuracao
End Sub
Private Sub SetHHINI()
   Dim bOk As Boolean
   
   If mvarCalEvent Is Nothing Then Exit Sub
   If Not mvarMe.CmbHHINI.Visible Then Exit Sub
   With mvarMe
      bOk = IsNull(.CmbHHINI.Value)
      bOk = bOk Or (Format(mvarCalEvent.StartTime, "hh:mm") = .TxtHHINI.Text)
      bOk = bOk And IsNull(.CmbHHFIM.Value)
      bOk = bOk And Not (bAtendFechado Or bVendaFechada)
      If bOk Then
         .TxtHHINI.Text = Format(Sys.xDb.Sysdate(2), "hh:mm")
         .CmbHHINI.Value = Format(.TxtHHINI.Text, "hh:mm")
         .TxtHHINI.Enabled = Not IsNull(.CmbHHINI.Value)
         If .CmbHHINI.Enabled And .CmbHHINI.Visible Then
            .CmbHHINI.SetFocus
         End If
      End If
      If Trim(.TxtHHFIM.Text) = "" Or Trim(.TxtHHFIM.Text) = "" Then
         Call ExibirAviso("Hora Inicio ou Fim inválida.", "Atendimento")
      Else
         If CDate(.TxtHHFIM.Text) < CDate(.TxtHHINI.Text) Or IsNull(.CmbHHFIM.Value) Then
            .TxtHHFIM.Text = Format(DateAdd("n", DateDiff("n", mvarCalEvent.StartTime, mvarCalEvent.EndTime), CDate(.TxtHHINI.Text)), "hh:mm")
            .CmbHHFIM.Value = Format(.TxtHHFIM.Text, "hh:mm")
         End If
      End If
   End With
End Sub
Private Sub SetHHFIM()
   With mvarMe
      If IsNull(.CmbHHFIM.Value) Then
         .TxtHHFIM.Text = Format(Sys.xDb.Sysdate(2), "hh:mm")
         .CmbHHFIM.Value = Format(.TxtHHFIM.Text, "hh:mm")
         .TxtHHFIM.Enabled = Not IsNull(.CmbHHFIM.Value)
         If .CmbHHFIM.Enabled And .CmbHHFIM.Visible Then
            .CmbHHFIM.SetFocus
         End If
      End If
   End With
End Sub

Private Sub mvarMe_CmbIDFUNCIONARIOClick()
   Call SetHHINI
End Sub

Private Sub mvarMe_CmbIDSALAClick()
   If mvarMe.CmbIDSALA.ListIndex >= 0 And mvarMe.CmbIDSALA.Visible Then
      IDSALA = mvarMe.CmbIDSALA.ItemData(mvarMe.CmbIDSALA.ListIndex)
      Call MontaCmbMaquina
      If mvarIDATENDIMENTO <> 0 Then
         Call LocalizarCombo(mvarMe.CmbIDMAQUINA, CStr(mvarIDMAQUINA), True, True)
      End If
   End If
End Sub

Private Sub mvarMe_CmdCancelClick()
   On Error Resume Next
   'Call RefreshEvent(mvarSys, mvarCalControl, mvarCalEvent, False)
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdChaveClick()
   If SenhaMestre(mvarSys) Then
      Screen.MousePointer = vbHourglass
      bAtendFechado = False
      bVendaFechada = False
      If Not mvarCalEvent Is Nothing Then
         bVendaFechada = False
      End If
      Call PopulaTela
      mvarMe.CmdFechar.Enabled = False
      bChave = True
      Screen.MousePointer = vbDefault
   Else
      Call ExibirInformacao("Senha Inválida")
   End If
End Sub
Private Sub mvarMe_CmdEditarClick()
   Dim lRow As Long
   Dim lCol As Long
   
   On Error Resume Next
   mvarMe.MousePointer = vbHourglass
   
   mvarMe.CmdEditar.Enabled = False
   If Not bExisteVenda Then bChave = True
   
   lRow = mvarMe.GrdVenda.CurRow
   lCol = mvarMe.GrdVenda.CurCol

   If xVal(mvarMe.GrdVenda.CellValue(lRow, "IDVENDA")) > 0 Then
      Call ExibirVenda(mvarMe.GrdVenda.CellValue(lRow, "IDVENDA"))
      Call PopulaGRDVENDA
   End If
   If lRow <= mvarMe.GrdVenda.RowCount - 1 And lRow > 0 Then
       Call mvarMe.GrdVenda.SetCurCell(lRow, lCol)
   End If
   
   mvarMe.CmdEditar.Enabled = True
   mvarMe.MousePointer = vbDefault
End Sub

Private Sub mvarMe_CmdFecharClick()
   Dim Sql As String
   
   Dim bSalvar As Boolean
   If VerificaAtendimento(bVerifVenda:=mvarMe.CmdChave.Visible, bExibeMsg:=Not bChave) Or bChave Then
      If SalvarAtendimento Then
         Sql = "Update OATENDIMENTO "
         Sql = Sql & " Set SITATEND=" & SqlStr(gSITATEND_Fechado)
         Sql = Sql & " , ALTERSTAMP=1"
         Sql = Sql & " , TIMESTAMP =GetDate()"
         Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
         Sql = Sql & " And IDATENDIMENTO=" & SqlNum(mvarIDATENDIMENTO)
         
         bAtendFechado = mvarSys.xDb.Executa(Sql)
         
         Set mvarMe.CmdFechar.Picture = mvarMe.lstImage.ListImages(IIf(bAtendFechado, 2, 1)).Picture
         mvarMe.CmdFechar.Enabled = Not bAtendFechado
         
         On Error Resume Next
         Call RefreshEvent(mvarSys, mvarCalControl, mvarCalEvent, False)
      End If
   End If
   
   
   
End Sub

Private Sub mvarMe_CmdIDCLIENTEClick()
   Call ShowCliente(mvarSys, mvarIDLOJA, mvarIDCLIENTE)
End Sub

Private Sub mvarMe_CmdLovClick()
   Dim MyCliente As Object

   If bVendaFechada Or bAtendFechado Then
      Call ExibirInformacao(IIf(bVendaFechada, "Venda fechada.", "Atendimento fechado."))
      Exit Sub
   End If
   
   Set MyCliente = CriarObjeto("Contato3R.TL_NovoLov", False)
   With MyCliente
      Set .Sys = mvarSys
      '.Filtro = Trim(mvarMe.TxtNOME.Text)
      Call .ShowClientes
      If .IdSelecionado <> 0 Then
         Call PopulaCliente(.IdSelecionado)
      End If
   End With
   Set MyCliente = Nothing
End Sub

Private Sub mvarMe_CmdMaximizedClick()
   Dim MyEditor As Object
   
   Set MyEditor = CriarObjeto("DSACTIVE.EDITOR")
   With MyEditor
      .FundoTela = Empty
      .Texto = mvarMe.TxtObsOutras.Text
      
      .Caption = "Observações"
      .Width = 8000 'mvarMe.Width
      .Height = 7000 'mvarMe.Height
      .Edicao = False
      .SetIcon mvarMe

      .Show
      
      'ExibeEditor = .Texto
   End With
   Set MyEditor = Nothing
End Sub

Private Sub mvarMe_CmdNovoClick()
   Dim bExiste As Boolean
   Dim nSaldo As Integer
   
   bExiste = VerificarPacote(True, nSaldo)
   If nSaldo > 0 Or Not bExiste Then
      Call ExibirVenda(-1)
      Call PopulaGRDVENDA
   End If
End Sub

Private Sub mvarMe_CmdOkClick()
   Call mvarMe_CmdSalvarClick
   If mvarIDATENDIMENTO <> 0 Then
      Unload mvarMe
   End If
End Sub
Private Function VerificaAtendimento(Optional bVerifVenda As Boolean = True, Optional bDelete As Boolean, Optional bExibeMsg As Boolean = True, Optional bParaVenda As Boolean = False) As Boolean
   Dim nAux As Integer
   
   If Trim(GetTag(mvarMe.TxtNOME, "TEL1", "") & GetTag(mvarMe.TxtNOME, "TEL2", "") & GetTag(mvarMe.TxtNOME, "EMAIL", "")) = "" Then
      Call ExibirInformacao("Informe um contato do cliente.")
      mvarMe.TxtTEL1.SetFocus
      Exit Function
   End If
   If Trim(mvarMe.TxtNOME.Text) = "" Then
      Call ExibirInformacao("Informe o nome do cliente.")
      mvarMe.TxtNOME.SetFocus
      Exit Function
   End If
   If mvarIDCLIENTE = 0 Then
      Call ExibirInformacao("Cliente sem identificação." & vbNewLine & "Favor verificar cadastro.")
      mvarMe.TxtTEL1.SetFocus
      Exit Function
   End If
   With mvarMe.GrdSESSAO
      If .RowCount <= 1 Then
         Call ExibirInformacao("Itens de atendimento inválidos.")
         .SetFocus
         Exit Function
      Else
         If Val(.CellValue(1, "IDTPSERVICO")) = 0 Then
            Call ExibirInformacao("Serviço inválido.")
            Call .SetCurCell(1, "IDTPSERVICO")
            .SetFocus
            Exit Function
         End If
         If Val(.CellValue(1, "IDTPTRATAMENTO")) = 0 Then
            
            Call ExibirInformacao("Tratamento inválido.")
            Call .SetCurCell(1, "IDTPTRATAMENTO")
            .SetFocus
            Exit Function
         End If
         If Val(.CellValue(1, "IDAREA")) = 0 And .ColVisible("IDAREA") And Mid(.CellText(1, "IDAREA"), 1, 1) <> "-" Then
         'If Val(.CellValue(1, "IDAREA")) = 0 Then
            Call ExibirInformacao("Área inválida.")
            Call .SetCurCell(1, "IDAREA")
            .SetFocus
            Exit Function
         End If
         If bIsFOTO And mvarSys.Propriedades("TPDIR") = "1" Then
            If Val(.CellValue(1, "IDDIRECAO")) = 0 Then
               Call ExibirInformacao("Direção inválida.")
               Call .SetCurCell(1, "IDDIRECAO")
               .SetFocus
               Exit Function
            End If
         End If
      End If
   End With
   With mvarMe.GrdPRODUTO
      If .RowCount > 1 Then
         If Val(.CellValue(1, "QTDCOMPRA")) = 0 Then
            Call ExibirInformacao("Quantidade de produto inválida.")
            Call .SetCurCell(1, "QTDCOMPRA")
            mvarMe.TabControl2.SelectedItem = 2
            If .Enabled And .Visible Then
               .SetFocus
            End If
            Exit Function
         End If
      End If
   End With
   If Trim(mvarMe.CmbIDFUNCIONARIO.Text) = "" Or mvarMe.CmbIDFUNCIONARIO.ListIndex = -1 Then
      Call ExibirInformacao("Informe a operadora.")
      mvarMe.CmbIDFUNCIONARIO.Enabled = True
      mvarMe.CmbIDFUNCIONARIO.SetFocus
      Exit Function
   End If
   If Trim(mvarMe.CmbIDSALA.Text) = "" Or mvarMe.CmbIDSALA.ListIndex = -1 Then
      Call ExibirInformacao("Informe a sala de operação.")
      mvarMe.CmbIDSALA.Enabled = True
      mvarMe.CmbIDSALA.SetFocus
      Exit Function
   End If
   If (Trim(mvarMe.CmbIDMAQUINA.Text) = "" Or mvarMe.CmbIDMAQUINA.ListIndex = -1) And bFLGDISPARO Then
      Call ExibirInformacao("Informe a máquina de operação.")
      mvarMe.CmbIDMAQUINA.Enabled = True
      mvarMe.CmbIDMAQUINA.SetFocus
      Exit Function
   End If
   If bParaVenda Then
      If IsNull(mvarMe.CmbHHINI.Value) Or IsNull(mvarMe.CmbHHFIM.Value) Then
         Call ExibirInformacao("Para gerar venda é necessário finalizar tempo da sessão.")
         If IsNull(mvarMe.CmbHHINI.Value) Then
            mvarMe.CmbHHINI.CheckBox = True
            mvarMe.CmbHHINI.SetFocus
         ElseIf IsNull(mvarMe.CmbHHFIM.Value) Then
            mvarMe.CmbHHFIM.CheckBox = True
            mvarMe.TxtHHFIM.Enabled = True
            mvarMe.TxtHHFIM.Text = Format(Sys.xDb.Sysdate(3), "hh:mm")
            mvarMe.CmbHHFIM.Value = mvarMe.TxtHHFIM.Text
            mvarMe.TxtHHFIM.SetFocus
         End If
         Exit Function
      Else
         If Not bExisteVenda Then
            If CDate(Format(mvarMe.CmbHHFIM.Value, "hh:mm")) > CDate(mvarSys.xDb.Sysdate(2)) Then
               If mvarMe.CmbHHFIM.CheckBox = True Then
                  nAux = DateDiff("n", CDate(Format(mvarMe.CmbHHFIM.Value, "hh:mm")), CDate(mvarSys.xDb.Sysdate(2)))
                  If nAux >= 5 And nAux < 60 Then
                     If ExibirPergunta("Atualiza hora do atendimento?", "Hora de Fim") = vbYes Then
                        mvarMe.TxtHHFIM.Text = Format(Sys.xDb.Sysdate(2), "hh:mm")
                        mvarMe.CmbHHFIM.Value = mvarMe.TxtHHFIM.Text
                     End If
                  End If
               Else
                  mvarMe.TxtHHFIM.Text = Format(Sys.xDb.Sysdate(2), "hh:mm")
                  mvarMe.CmbHHFIM.Value = mvarMe.TxtHHFIM.Text
               End If
            End If
            If CDate(Format(mvarMe.CmbHHFIM.Value, "dd/mm/yyyy")) = CDate(mvarSys.xDb.Sysdate(1)) Then
               nAux = DateDiff("n", CDate(Format(mvarMe.CmbHHFIM.Value, "hh:mm")), CDate(mvarSys.xDb.Sysdate(2)))
               If nAux >= 5 And nAux < 60 Then
                  If ExibirPergunta("Atualiza hora do atendimento?", "Hora de Fim") = vbYes Then
                     mvarMe.TxtHHFIM.Text = Format(Sys.xDb.Sysdate(2), "hh:mm")
                     mvarMe.CmbHHFIM.Value = mvarMe.TxtHHFIM.Text
                  End If
               End If
            End If
         End If
      End If
   End If

   If bVerifVenda Then
      If bAtendFechado Then
         If bExibeMsg Then ExibirAviso ("Atendimento fechado e não pode ser alterado.")
         GoTo Saida
      End If
      If bExisteVenda Then
         If bVendaFechada Then
            If bExibeMsg Then ExibirAviso ("Atendimento possui venda fechada e não pode ser alterado.")
            GoTo Saida
         ElseIf bDelete Then
            Call ExibirAviso("Atendimento possui Venda e não pode ser alterado.")
            GoTo Saida
         End If
      End If
   End If
   VerificaAtendimento = True
Saida:
End Function
Private Function SalvarAtendimento(Optional pSalvaEvento As Boolean = True, Optional pSalvaTarefa As Boolean = True) As Boolean
   Dim i          As Integer
   Dim cQueries   As Collection
   Dim bResult    As Boolean
   Dim TbSessao   As Object 'TB_OSESSAO
   Dim TbProd     As Object 'TB_OATENDIMENTO_PRODUTO
   Dim bAchou     As Boolean
   Dim bisDirt    As Boolean
   
   Dim NgCal      As NG_Calendario
   Dim sDate      As String
   Dim sDateIni   As String
   Dim sDateFim   As String
   Dim sTbDateIni As String
   Dim sTbDateFim As String
               
   If mvarIDATENDIMENTO = 0 Then mvarIDATENDIMENTO = -1
               
   '* Atendimento
   If Not IsNull(mvarMe.CmbHHINI.Value) And Not IsNull(mvarMe.CmbHHFIM.Value) Then
      sDate = Format(mvarMe.CmbDTATEND, "dd/mm/yyyy")
      If mvarSys.xDb.Sysdate() <> sDate Then
      
         sDateIni = sDate & Format(mvarMe.TxtHHINI.Text, " hh:mm")
         sDateFim = sDate & Format(mvarMe.TxtHHFIM.Text, " hh:mm")
         
         sTbDateIni = IIf(mvarTbAtend.HHINI = "", "01/01/1600 00:00", mvarTbAtend.HHINI)
         sTbDateFim = IIf(mvarTbAtend.HHFIM = "", "01/01/1600 00:00", mvarTbAtend.HHFIM)
   
         If CDate(sTbDateIni) <> CDate(sDateIni) Or CDate(sTbDateFim) <> CDate(sDateFim) Then
            Set NgCal = New NG_Calendario
            Set NgCal.Sys = mvarSys
            If NgCal.VerificaSobreposicao(mvarIDEVENTO, mvarIDSALA, sDateIni, sDateFim) Then
               'GoTo Saida
            End If
         End If
      End If
   End If
   bisDirt = False
   If PopulaClasse_Atendimento(mvarTbAtend) Then
      If mvarTbAtend.isDirt Then
         bResult = mvarTbAtend.Salvar
         bisDirt = True
      Else
         bResult = True
      End If
   Else
      GoTo Saida
   End If
   
   If bResult Then
      mvarIDATENDIMENTO = mvarTbAtend.IDATENDIMENTO
      mvarMe.Caption = "Atendimento Nº" & StrZero(mvarIDATENDIMENTO, mvarSys.Propriedades("LENIDVENDA"))
      '********
      '* Evento
      If mvarIDEVENTO <> 0 Then
         If bisDirt Then
            Call SalvarEvento(mvarTbAtend)
         End If
      End If
      '********
      '* Venda
      'Call SalvarVinculoVenda
                  
      Set cQueries = New Collection
      '* Itens
      If GetTag(mvarMe.GrdSESSAO, "ISDIRT", 0) = 1 Then
         If IsDirtSessao Then
            Set TbSessao = CriarObjeto("BANCO_3R.TB_OSESSAO", False)
            Set TbSessao.xDb = mvarSys.xDb
            TbSessao.IDLOJA = mvarIDLOJA
            TbSessao.IDATENDIMENTO = mvarIDATENDIMENTO
      
            cQueries.Add TbSessao.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO, Ch_Where:="IDSESSAO >= " & mvarMe.GrdSESSAO.RowCount)
            With mvarMe.GrdSESSAO
               For i = 1 To .RowCount - 1
                  If Val(.CellValue(i, "IDSESSAO")) <> 0 And _
                     Val(.CellValue(i, "IDTPSERVICO")) <> 0 And _
                     Val(.CellValue(i, "IDTPTRATAMENTO")) <> 0 And _
                     (Val(.CellValue(i, "IDAREA")) <> 0 Or Not .ColVisible("IDAREA") Or Mid(.CellText(i, "IDAREA"), 1, 1) = "-") Then
                     
                     bAchou = TbSessao.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO, Ch_IDSESSAO:=Val(.CellValue(i, "IDSESSAO")))
                     TbSessao.IDLOJA = mvarIDLOJA
                     TbSessao.IDATENDIMENTO = mvarIDATENDIMENTO
                     TbSessao.IDSESSAO = Val(.CellValue(i, "IDSESSAO"))
   
                     If bIsFOTO Then
                        If TbSessao.DISPAROS <> Val(.CellValue(i, "DISPAROS")) And bAchou Then
                           Dim TbDisp As Object 'TB_OMAQDISPAROS
                           Dim Sql     As String
                           
                           Set TbDisp = CriarObjeto("BANCO_3R.TB_OMAQDISPAROS")
                           Set TbDisp.xDb = mvarSys.xDb
                           Sql = " IDATENDIMENTO=" & SqlNum(mvarIDATENDIMENTO)
                           Sql = Sql & " And IDSESSAO=" & SqlNum(TbSessao.IDSESSAO)
                           If TbDisp.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDMAQUINA:=mvarMe.CmbIDMAQUINA.ItemData(mvarMe.CmbIDMAQUINA.ListIndex), Ch_Where:=Sql) Then
                              TbSessao.DISPAROS = TbDisp.QTDDISPARO
                              ExibirInformacao "Sessão relacionada ao registro de disparos da máquina '" & mvarMe.CmbIDMAQUINA.Text & "'." & vbNewLine & vbNewLine & "Quantidade de disparos não pode ser alterada pelo atendimento."
                           Else
                              TbSessao.DISPAROS = Val(.CellValue(i, "DISPAROS"))
                           End If
                        Else
                           TbSessao.DISPAROS = Val(.CellValue(i, "DISPAROS"))
                        End If
                        TbSessao.FOTOTIPO = Val(.CellValue(i, "FOTOTIPO"))
                        TbSessao.IDDIRECAO = Val(.CellValue(i, "IDDIRECAO"))
                        TbSessao.POTENCIA = Val(.CellValue(i, "POTENCIA"))
                        TbSessao.TEMPOSESSAO = Val(.CellValue(i, "TEMPOSESSAO"))
                        'TbSessao.IDTPMANIPULO = Val(.CellValue(i, "IDTPMANIPULO"))
                        TbSessao.IDMANIPULO = Val(.CellValue(i, "IDMANIPULO"))
                        If TbSessao.IDMANIPULO = 0 Then TbSessao.IDMANIPULO = 1
                     End If
                     TbSessao.IDAREA = Val(.CellValue(i, "IDAREA"))
                     TbSessao.IDTPSERVICO = Val(.CellValue(i, "IDTPSERVICO"))
                     TbSessao.IDTPTRATAMENTO = Val(.CellValue(i, "IDTPTRATAMENTO"))
                     TbSessao.NUMSESSAO = Val(.CellValue(i, "NUMSESSAO"))
                     
                     cQueries.Add TbSessao.QrySave
                  End If
               Next
            End With
         End If
      End If
      
      '* Produtos
      If GetTag(mvarMe.GrdPRODUTO, "ISDIRT", 0) = 1 Then
         If IsDirtProduto Then
            Set TbProd = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_PRODUTO")
            Set TbProd.xDb = mvarSys.xDb
            TbProd.IDLOJA = mvarIDLOJA
            TbProd.IDATENDIMENTO = mvarIDATENDIMENTO
      
            cQueries.Add TbProd.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO)
            With mvarMe.GrdPRODUTO
               For i = 1 To .RowCount - 1
                  If Val(.CellValue(i, "IDPROD")) <> 0 Then
                     TbProd.IDLOJA = mvarIDLOJA
                     TbProd.IDATENDIMENTO = mvarIDATENDIMENTO
                     
                     TbProd.IDPROD = Val(.CellValue(i, "IDPROD"))
                     TbProd.QTDCOMPRA = Val(.CellValue(i, "QTDCOMPRA"))
                                       
                     cQueries.Add TbProd.QrySave
                  End If
               Next
            End With
         End If
      End If
      
      '* Verificar SITATEND
      If mvarTbAtend.SITATEND = "" Then mvarTbAtend.SITATEND = gSITATEND_Aberto
      If mvarTbAtend.SITATEND = gSITATEND_Fechado Then
         Call CarregaFechamento
         If Not bExisteVenda Or Not bVendaFechada Then
            mvarTbAtend.SITATEND = gSITATEND_Aberto
            If mvarTbAtend.isDirt Then
               bResult = bResult And mvarTbAtend.Salvar
            End If
         End If
      End If
      
      SalvarAtendimento = bResult
      If cQueries.Count > 0 Then
         If mvarSys.xDb.Executa(cQueries) Then
            SalvarAtendimento = True
            Call RefreshEvent(mvarSys, mvarCalControl, mvarCalEvent, False)
         Else
            Call ExibirInformacao("Erro ao gravar itens e produtos.")
         End If
      End If
   Else
      Call ExibirInformacao("Erro ao gravar atendimento.")
   End If
        
Saida:
   Set TbSessao = Nothing
   Set TbProd = Nothing
   Set cQueries = Nothing
End Function
Private Function IsDirtSessao() As Boolean
   Dim Sql As String
   Dim MyRS As Object
   Dim iCol As Integer
   Dim iRow As Integer
   Dim sCampo As String
   
   Sql = "Select *"
   Sql = Sql & " From OSESSAO"
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And IDATENDIMENTO=" & mvarIDATENDIMENTO
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      If MyRS.RecordCount = (mvarMe.GrdSESSAO.RowCount - 1) Then
         iRow = 0
         While Not MyRS.EOF
            iRow = iRow + 1
            For iCol = 1 To mvarMe.GrdSESSAO.ColCount
               sCampo = mvarMe.GrdSESSAO.ColKey(iCol)
               If MyRS(sCampo) <> mvarMe.GrdSESSAO.CellValue(iRow, iCol) Or IsNull(MyRS(sCampo)) Then
                  If MyRS(sCampo) = 0 Or IsNull(MyRS(sCampo)) Then
                     If mvarMe.GrdSESSAO.CellValue(iRow, iCol) <> "" Then
                        IsDirtSessao = True
                        GoTo Saida
                     End If
                  Else
                     IsDirtSessao = True
                     GoTo Saida
                  End If
               End If
            Next
            MyRS.MoveNext
         Wend
      Else
         IsDirtSessao = True
      End If
   Else
      IsDirtSessao = (mvarMe.GrdSESSAO.RowCount > 1)
   End If
Saida:
   IsDirtSessao = IsDirtSessao
End Function
Private Function IsDirtProduto() As Boolean
   Dim Sql As String
   Dim MyRS As Object
   Dim iCol As Integer
   Dim iRow As Integer
   Dim sCampo As String
   
   Sql = ""
   Sql = Sql & " Select RANK() OVER(ORDER BY A.IDPROD) ITEM, A.IDATENDIMENTO"
   Sql = Sql & ", A.IDLOJA, S.IDPROD, S.NMPROD, A.QTDCOMPRA, S.UNIDCONTROLE" & vbNewLine
   Sql = Sql & " From OATENDIMENTO_PRODUTO A" & vbNewLine
   Sql = Sql & " Join SPRODUTO S On S.IDLOJA=A.IDLOJA AND S.IDPROD=A.IDPROD" & vbNewLine
   Sql = Sql & " Where A.IDLOJA=" & mvarIDLOJA & vbNewLine
   Sql = Sql & " And A.IDATENDIMENTO=" & mvarIDATENDIMENTO & vbNewLine
   Sql = Sql & " Order By A.IDPROD"
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      If MyRS.RecordCount = (mvarMe.GrdPRODUTO.RowCount - 1) Then
         iRow = 0
         While Not MyRS.EOF
            iRow = iRow + 1
            For iCol = 1 To mvarMe.GrdPRODUTO.ColCount
               sCampo = mvarMe.GrdPRODUTO.ColKey(iCol)
               If MyRS(sCampo) <> mvarMe.GrdPRODUTO.CellValue(iRow, iCol) Then
                  If MyRS(sCampo) = 0 Then
                     If mvarMe.GrdPRODUTO.CellValue(iRow, iCol) <> "" Then
                        IsDirtProduto = True
                        GoTo Saida
                     End If
                  Else
                     IsDirtProduto = True
                     GoTo Saida
                  End If
               End If
            Next
            MyRS.MoveNext
         Wend
      Else
         IsDirtProduto = True
      End If
   Else
      IsDirtProduto = (mvarMe.GrdPRODUTO.RowCount > 1)
   End If
Saida:
   IsDirtProduto = IsDirtProduto
End Function
'Private Sub RefreshEvent()
'   If mvarIDATENDIMENTO <> 0 Then
'      If Not mvarCalEvent Is Nothing And Not mvarCalControl Is Nothing Then
'         If mvarCalEvent.Id > 0 Then
'            mvarCalEvent.CustomProperties("EXIBEMSG") = False
'            mvarCalControl.DataProvider.ChangeEvent mvarCalEvent
'            mvarCalControl.Enabled = True
'            mvarCalControl.RedrawControl
'            mvarCalControl.Populate
'            mvarCalEvent.CustomProperties("EXIBEMSG") = True
'         End If
'      End If
'   End If
'End Sub
Private Function SalvarEvento(Optional TBAtend As Object) As Boolean
   Dim i          As Integer
   Dim TbEvt      As Object 'TB_OEVENTOAGENDA
   Dim TbServ     As Object 'TB_OSERVICOEVT
   Dim cQueries   As Collection
   Dim bOk        As Boolean
   Dim NgCal      As NG_Calendario
   Dim nLabelID   As Long
   Dim bAchouEvt  As Boolean
   Dim sDate      As String
   Dim bTarefas   As Boolean
   Dim bisDirt    As Boolean
   
   If mvarCalEvent Is Nothing Then
      If Not mvarCalControl Is Nothing Then
         mvarCalEvent = mvarCalControl.DataProvider.GetEvent(mvarIDEVENTO)
      End If
   End If

   bisDirt = False
   If mvarCalEvent Is Nothing Then
      bisDirt = True
   Else
      bisDirt = bisDirt Or (xVal(mvarCalEvent.CustomProperties("IDCLIENTE")) <> mvarIDCLIENTE)
      If Not IsNull(mvarMe.CmbHHINI.Value) And Not bisDirt Then
         bisDirt = bisDirt Or (Format(mvarCalEvent.StartTime, "hh:mm") <> Format(mvarMe.CmbHHINI.Value, "hh:mm"))
      End If
      If bVendaFechada Then
         bisDirt = bisDirt Or (mvarCalEvent.CustomProperties("VENDAFECHADA") <> bVendaFechada)
      End If
      If Not IsNull(mvarMe.CmbHHFIM.Value) And Not bisDirt Then
         bisDirt = bisDirt Or (Format(mvarCalEvent.EndTime, "hh:mm") <> Format(mvarMe.CmbHHFIM.Value, "hh:mm"))
      End If
      If Not IsNull(mvarMe.CmbDTATEND.Value) And Not bisDirt Then
         bisDirt = bisDirt Or (Format(mvarCalEvent.StartTime, "dd/mm/yyyy") <> Format(mvarMe.CmbDTATEND.Value, "dd/mm/yyyy"))
      End If
      If Not IsNull(mvarMe.CmbDTATEND.Value) And Not bisDirt Then
         bisDirt = bisDirt Or (Format(mvarCalEvent.EndTime, "dd/mm/yyyy") <> Format(mvarMe.CmbDTATEND.Value, "dd/mm/yyyy"))
      End If
      If mvarMe.GrdSESSAO.CellValue(1, "IDTPSERVICO") > 0 And Not bisDirt Then
         bisDirt = bisDirt Or (mvarCalEvent.BusyStatus <> DefineShowAs(mvarMe.GrdSESSAO.CellValue(1, "IDTPSERVICO")))
      End If
      If mvarMe.GrdSESSAO.CellValue(1, "IDTPTRATAMENTO") > 0 And Not bisDirt Then
         bisDirt = bisDirt Or (mvarCalEvent.Label <> DefineLabelID(mvarMe.GrdSESSAO.CellValue(1, "IDTPTRATAMENTO"), mvarMe.GrdSESSAO.CellValue(1, "IDTPSERVICO")))
      End If
      If mvarMe.CmbIDSALA.ListIndex >= 0 And Not bisDirt Then
         bisDirt = bisDirt Or (mvarCalEvent.ScheduleID <> (mvarIDLOJA * 1000) + Val(mvarMe.CmbIDSALA.ItemData(mvarMe.CmbIDSALA.ListIndex)))
      End If
   End If
   
   If Not bisDirt Then
      SalvarEvento = True
      GoTo Saida
   End If
   
   Set TbEvt = CriarObjeto("BANCO_3R.TB_OEVENTOAGENDA", False)
   Set TbEvt.xDb = mvarSys.xDb
   bAchouEvt = TbEvt.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEVENTO:=mvarIDEVENTO)
   bOk = True
   If bAchouEvt Then
      If Not IsNull(mvarMe.CmbHHINI.Value) Then
         sDate = Format(mvarMe.CmbDTATEND.Value, "dd/mm/yyyy")
         TbEvt.StartDateTime = sDate & " " & Format(mvarMe.TxtHHINI.Text, "hh:mm:00")
         If IsNull(mvarMe.CmbHHFIM.Value) Then
            TbEvt.EndDateTime = DateAdd("n", 30, CDate(TbEvt.StartDateTime))
         Else
            TbEvt.EndDateTime = sDate & " " & Format(mvarMe.TxtHHFIM.Text, "hh:mm:00")
         End If
      End If
      TbEvt.Subject = mvarMe.TxtNOME & " - " & mvarMe.TxtTEL1
      TbEvt.IDCLIENTE = mvarIDCLIENTE
      TbEvt.IDSALA = mvarMe.CmbIDSALA.ItemData(mvarMe.CmbIDSALA.ListIndex)
      
      If mvarMe.GrdSESSAO.CellValue(1, "IDTPSERVICO") > 0 Then
         TbEvt.BusyStatus = DefineShowAs(mvarMe.GrdSESSAO.CellValue(1, "IDTPSERVICO"))
      End If
      If mvarMe.GrdSESSAO.CellValue(1, "IDTPTRATAMENTO") > 0 Then
         TbEvt.LabelID = DefineLabelID(mvarMe.GrdSESSAO.CellValue(1, "IDTPTRATAMENTO"), mvarMe.GrdSESSAO.CellValue(1, "IDTPSERVICO"))
      End If
      TbEvt.ScheduleID = mvarSys.Objetos("LOJAS")(CStr(mvarIDLOJA)).Salas(CStr(TbEvt.IDSALA)).ScheduleID
      If Not mvarCalEvent Is Nothing Then
         TbEvt.CustomPropertiesXMLData = mvarCalEvent.CustomProperties.SaveToString
      End If
      
      bTarefas = TbEvt.isDirt
      If TbEvt.isDirt Then bOk = TbEvt.Salvar
   End If
   
   Set NgCal = New NG_Calendario
   Set NgCal.Sys = mvarSys
   If Not bChave Then
      If Not mvarCalEvent Is Nothing Then
         If mvarCalEvent.Id <> mvarIDEVENTO Then
            mvarCalEvent = mvarCalControl.DataProvider.GetEvent(mvarIDEVENTO)
         End If
         mvarCalEvent.ScheduleID = TbEvt.ScheduleID
         mvarCalEvent.CustomProperties("IDSALA") = TbEvt.IDSALA
         
         If Not NgCal.VerificaEvento(mvarCalEvent, False, False) Then GoTo Saida
      End If
   End If

   If bAchouEvt Then
      If bOk Then
         Set cQueries = New Collection
         
         '* Salvar Tarefas
         Dim NgTarefa As Object
         Dim bTeste  As Boolean
         bTarefas = bTarefas And Not (TBAtend Is Nothing)
         If bTarefas Then
            bTeste = (1 = Val(mvarMe.GrdSESSAO.CellValue(1, "IDTPSERVICO")))
            Set NgTarefa = CriarObjeto("Tarefa3R.NG_TAREFA", False)
            Set NgTarefa.Sys = mvarSys
            Call NgTarefa.SalvarTarefas(TbEvt, cQueries, bTeste)
         End If
         Set NgTarefa = Nothing
            
         
         '* Salvar Serviços
         Set TbServ = CriarObjeto("BANCO_3R.TB_OSERVICOEVT", False)
         With TbServ
            Set .xDb = mvarSys.xDb
            If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEVENTO:=mvarIDEVENTO, Ch_IDSERVICOEVT:=1) Then
               TbServ.IDTPSERVICO = Val(mvarMe.GrdSESSAO.CellValue(1, "IDTPSERVICO"))
               TbServ.IDTPTRATAMENTO = Val(mvarMe.GrdSESSAO.CellValue(1, "IDTPTRATAMENTO"))
               If TbServ.isDirt Then
                  cQueries.Add TbServ.QrySave
               End If
            
            Else
               .IDLOJA = mvarIDLOJA
               .IDEVENTO = mvarIDEVENTO
               With mvarMe.GrdSESSAO
                  For i = 1 To .RowCount - 1
                     If Val(.CellValue(i, "IDSESSAO")) <> 0 And _
                     Val(.CellValue(i, "IDTPSERVICO")) <> 0 And _
                     Val(.CellValue(i, "IDTPTRATAMENTO")) <> 0 And _
                     (Val(.CellValue(i, "IDAREA")) <> 0 Or Not .ColVisible("IDAREA")) Then
                        TbServ.IDSERVICOEVT = Val(.CellValue(i, "IDSESSAO"))
                        TbServ.IDTPSERVICO = Val(.CellValue(i, "IDTPSERVICO"))
                        TbServ.IDTPTRATAMENTO = Val(.CellValue(i, "IDTPTRATAMENTO"))
                        TbServ.IDAREA = Val(.CellValue(i, "IDAREA"))
                        
                        cQueries.Add TbServ.QrySave
                     End If
                  Next
               End With
            End If
            If cQueries.Count > 0 Then
               bOk = mvarSys.xDb.Executa(cQueries)
            End If
         End With
      End If
      
      If bOk Then
         If Not mvarCalEvent Is Nothing Then
            On Error Resume Next
            
            mvarCalEvent.Subject = TbEvt.Subject
            mvarCalEvent.StartTime = TbEvt.StartDateTime
            mvarCalEvent.EndTime = TbEvt.EndDateTime
            mvarCalEvent.BusyStatus = TbEvt.BusyStatus
            mvarCalEvent.Label = TbEvt.LabelID
            
            
            '***************
            '* Propriedades Customizadas
            mvarCalEvent.CustomProperties("EXIBEMSG") = bAtendFechado
            Call SetCustomEvent(mvarCalEvent, mvarSys, pFLGCONFIRMADO:=TbEvt.FLGCONFIRMADO, pFLGCANCELADO:=TbEvt.FLGCANCELADO, pIDATENDIMENTO:=mvarIDATENDIMENTO, pIDCLIENTE:=mvarIDCLIENTE, pIDSALA:=TbEvt.IDSALA)
            Call mvarCalControl.DataProvider.ChangeEvent(mvarCalEvent)
            mvarCalControl.RedrawControl
            mvarCalControl.Populate
         End If
      End If
      
   End If
   SalvarEvento = True
   
Saida:
   Set TbEvt = Nothing
   Set TbServ = Nothing
   Set cQueries = Nothing
End Function
Private Function SalvarVinculoVenda() As Boolean
   Dim MyTb    As Object 'TB_OATENDIMENTO_VENDA
   Dim MyTbV   As Object 'TB_OATENDIMENTO_VENDA
   Dim bExiste As Boolean
   Dim Queries As Collection
   Dim Sql As String
   Dim sIn As String
   Dim i As Integer
   Dim MyVenda As Object 'TB_CVENDA
         
   Set Queries = New Collection
   
   If GetTag(mvarMe.GrdVenda.Tag, "ISDIRT", 0) = 0 Then Exit Function
   If mvarIDATENDIMENTO = 0 Then Exit Function
   
   Set MyTb = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_VENDA")
   Set MyTb.xDb = mvarSys.xDb
   With mvarMe.GrdVenda
      sIn = ""
      For i = 1 To .RowCount - 1
         bExiste = MyTb.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=xVal(.CellValue(i, "IDVENDA")), Ch_IDATENDIMENTO:=mvarIDATENDIMENTO)
         If Not bExiste Then
            MyTb.IDATENDIMENTO = mvarIDATENDIMENTO
            MyTb.IDLOJA = mvarIDLOJA
            MyTb.IDVENDA = xVal(.CellValue(i, "IDVENDA"))
            Queries.Add MyTb.QryInsert
         End If
         sIn = sIn & "|" & xVal(.CellValue(i, "IDVENDA")) & "|"
      Next
      If MyTb.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO) Then
         While Not MyTb.Rs.EOF
            If InStr(sIn, "|" & MyTb.Rs("IDVENDA") & "|") = 0 Then
               Set MyTbV = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_VENDA")
               Set MyTbV.xDb = mvarSys.xDb

               Sql = "UPDATE OSESSAO"
               Sql = Sql & " Set IDVENDA=0"
               Sql = Sql & " , IDITEM=0"
               Sql = Sql & " , ALTERSTAMP=1"
               Sql = Sql & " , TIMESTAMP=GETDATE()"
               Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
               Sql = Sql & " And IDATENDIMENTO=" & mvarIDATENDIMENTO
               Sql = Sql & " And IDVENDA =" & SqlNum(MyTb.Rs("IDVENDA") & "")
               Queries.Add Sql

               Queries.Add MyTbV.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=MyTb.Rs("IDVENDA"), Ch_IDATENDIMENTO:=mvarIDATENDIMENTO)
            End If
            MyTb.Rs.MoveNext
         Wend
      End If
      If Queries.Count > 0 Then
         SalvarVinculoVenda = False
         If mvarSys.xDb.Executa(Queries) Then
            SalvarVinculoVenda = True
            
            If MyTb.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO) Then
               Set MyVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
               Set MyVenda.xDb = mvarSys.xDb
               
               If MyVenda.Pesquisar(Ch_IDLOJA:=MyTb.IDLOJA, Ch_IDVENDA:=MyTb.IDVENDA) Then
                  '***************
                  '* Propriedades Customizadas
                  Call SetCustomEvent(mvarCalEvent, mvarSys, pIDVENDA:=MyVenda.IDVENDA, pSITVENDA:=MyVenda.SITVENDA)
                  'Call mvarCalControl.DataProvider.ChangeEvent(mvarCalEvent)
               End If
            End If
         End If
      End If
   End With
End Function
Private Function PopulaClasse_Atendimento(ByRef pTBAtend As Object) As Boolean
   Dim sDate   As String
   Dim sMsg    As String
   Dim lRow    As Integer
   
   If pTBAtend Is Nothing Then
      Set pTBAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
   End If
   
   With pTBAtend
      .xDb = mvarSys.xDb
      Call .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO)
      .IDATENDIMENTO = mvarIDATENDIMENTO
      If .IDCLIENTE <> 0 And .IDCLIENTE <> mvarIDCLIENTE Then
         sMsg = "Esta operação irá trocar o Cliente do evento."
         sMsg = sMsg & vbNewLine + vbNewLine
         sMsg = sMsg & "Deseja continuar?"
         If vbNo = ExibirPergunta(sMsg, "Atendimento") Then
            GoTo Saida
         Else
            Dim nNUMSESSAO As Long
            Dim nIDDIRECAO As Long
            Dim nFOTOTIPO   As Integer
            
            'Set TbSessao = CriarObjeto("BANCO_3R.TB_OSESSAO")
            'Set TbSessao.xdb = mvarSys.xdb
            'TbSessao.IDLOJA = mvarIDLOJA
            'TbSessao.IDATENDIMENTO = mvarIDATENDIMENTO
            'cQueries.Add TbSessao.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO)
            With mvarMe.GrdSESSAO
               For lRow = 1 To .RowCount - 1
                  Call SetNUMSESSAO(xVal(.CellValue(lRow, "IDTPSERVICO")), xVal(.CellValue(lRow, "IDTPTRATAMENTO")), xVal(.CellValue(lRow, "IDAREA")), nNUMSESSAO, nIDDIRECAO, nFOTOTIPO, pDTATEND:=pTBAtend.dtatend)
                  If .CellValue(lRow, "NUMSESSAO") <> nNUMSESSAO Or .CellValue(lRow, "IDDIRECAO") <> nIDDIRECAO Then
                     .CellValue(lRow, "NUMSESSAO") = nNUMSESSAO
                     .CellValue(lRow, "IDDIRECAO") = nIDDIRECAO
                     .CellValue(lRow, "FOTOTIPO") = nFOTOTIPO
                     Call SetTag(mvarMe.GrdSESSAO, "ISDIRT", 1)
                  End If
               Next
            End With
         End If
      End If
      
      .IDLOJA = mvarIDLOJA
      
      sDate = Format(mvarMe.CmbDTATEND.Value, "dd/mm/yyyy ")
      If IsNull(mvarMe.CmbHHFIM) Then
         .HHFIM = ""
      Else
         .HHFIM = sDate & Format(mvarMe.TxtHHFIM.Text, "hh:mm:00")
      End If
      If IsNull(mvarMe.CmbHHINI) Then
         .HHINI = ""
         .dtatend = mvarMe.CmbDTATEND.Value
      Else
         .HHINI = sDate & Format(mvarMe.TxtHHINI.Text, "hh:mm:00")
         .dtatend = .HHINI
      End If
      .IDCLIENTE = mvarIDCLIENTE
      .IDEVENTO = mvarIDEVENTO
      .IDFUNCIONARIO = mvarMe.CmbIDFUNCIONARIO.ItemData(mvarMe.CmbIDFUNCIONARIO.ListIndex)
      .IDMAQUINA = mvarMe.CmbIDMAQUINA.ItemData(mvarMe.CmbIDMAQUINA.ListIndex)
      .IDSALA = mvarMe.CmbIDSALA.ItemData(mvarMe.CmbIDSALA.ListIndex)
      .OBS = mvarMe.TxtOBS.Text
      '.SITATEND
      .IDMANIPULO = IIf(.IDMANIPULO = 0, 1, .IDMANIPULO)
   End With
   PopulaClasse_Atendimento = True
Saida:

End Function
Private Sub mvarMe_CmdSalvarClick()
   Screen.MousePointer = vbHourglass
   If VerificaAtendimento(bVerifVenda:=mvarMe.CmdChave.Visible, bExibeMsg:=Not bChave) Or bChave Then
      DoEvents
      If SalvarAtendimento Then
         On Error Resume Next
         Call ExibirResultado(mvarSys, True, 1)
         Call RefreshEvent(mvarSys, mvarCalControl, mvarCalEvent, False)
      Else
         Call ExibirResultado(mvarSys, False, 3)
      End If
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmdVendaClick()
   Dim lRow As Long
   Dim lCol As Long
   Dim nSaldo As Integer
   Dim bExiste  As Boolean
   Dim bMsg As Boolean
   Dim Sql As String
   Dim nAux As Integer
   
   mvarMe.MousePointer = vbHourglass
   
   mvarMe.CmdVenda.Enabled = False
   If Not bExisteVenda Then bChave = True
   
   lRow = mvarMe.GrdVenda.CurRow
   lCol = mvarMe.GrdVenda.CurCol
   
   bMsg = True
   If VerificaAtendimento(bVerifVenda:=False, bDelete:=False, bExibeMsg:=bMsg, bParaVenda:=True) Then
      bExiste = VerificarPacote(True, nSaldo)
      If nSaldo > 0 Or Not bExiste Then
         Call ExibirVenda
         Call PopulaGRDVENDA
      End If
      
      If lRow <= mvarMe.GrdVenda.RowCount - 1 And lRow > 0 Then
          Call mvarMe.GrdVenda.SetCurCell(lRow, lCol)
      End If
   End If
   mvarMe.CmdVenda.Enabled = True
   mvarMe.MousePointer = vbDefault
End Sub
Private Sub ExibirVenda(Optional pIDVENDA As Long)
   Dim TlVenda    As Object
   Dim TBAtend    As Object
   Dim NgFinanc   As Object
   Dim AtendSalvo As Boolean
   Dim Sql        As String
   Dim Msg        As String
   
   On Error GoTo Saida
   If Not mvarCalEvent Is Nothing Then
      mvarCalEvent.CustomProperties("EXIBIRVENDA") = True
      If bAtendFechado Then
         mvarCalEvent.CustomProperties("EXIBEMSG") = False
      End If
   End If
   If Not VerificaAtendimento(bVerifVenda:=False, bParaVenda:=True) Then Exit Sub
   AtendSalvo = SalvarAtendimento
   
   If AtendSalvo Then
      Set TlVenda = CriarObjeto("Financ3R.TL_Venda")
      With TlVenda
         Set .Sys = mvarSys
         Set .CalControl = mvarCalControl
         
         .IDLOJA = mvarIDLOJA
         .IDATENDIMENTO = mvarIDATENDIMENTO
         .IDCLIENTE = mvarIDCLIENTE
         If mvarMe.CmbIDFUNCIONARIO.ListIndex >= 0 Then
            .IDFUNCIONARIO = mvarMe.CmbIDFUNCIONARIO.ItemData(mvarMe.CmbIDFUNCIONARIO.ListIndex)
         End If
         
         If pIDVENDA <= 0 Then
            If mvarMe.GrdVenda.RowCount > 1 Then
               If pIDVENDA = 0 And xVal(mvarMe.GrdVenda.CellValue(mvarMe.GrdVenda.CurRow, "IDVENDA")) = 0 Then
                  pIDVENDA = -1
               End If
               '* Verifica se ainda há itens sem venda
               Sql = "Select * "
               Sql = Sql & " From OSESSAO"
               Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
               Sql = Sql & " And IDATENDIMENTO=" & mvarIDATENDIMENTO
               Sql = Sql & " And (isNull(IDITEM,0)=0 or isNull(IDVENDA,0)=0)"
               If mvarSys.xDb.ExisteReg(Sql) Then
                  pIDVENDA = -1
               End If
            
            Else
               pIDVENDA = -1
            End If
         End If
         If pIDVENDA = 0 Then
            If mvarMe.GrdVenda.CurRow > 0 Then
               .IDVENDA = xVal(mvarMe.GrdVenda.CellValue(mvarMe.GrdVenda.CurRow, "IDVENDA"))
            End If
            If mvarIDATENDIMENTO > 0 And .IDVENDA = 0 Then
               Sql = "Select * OSESSAO "
               Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
               Sql = Sql & " And IDATENDIMENTO=" & SqlNum(mvarIDATENDIMENTO)
               Sql = Sql & " And (isNull(IDVENDA)=0 or isNull(IDITEM)=0 )"
               If Not mvarSys.xDb.ExisteReg(Sql) Then
                  Exit Sub
               End If
            End If
            
         Else
            .IDVENDA = pIDVENDA
         End If
         
         .Show vbModal
         
         If .IDVENDA = 0 Then
            Dim lRow  As Integer
            With mvarMe.GrdVenda
               lRow = .CurRow
               If lRow <> .RowCount And lRow > 0 Then
                  .RemoveRow (.CurRow)
                  .Tag = SetTag(.Tag, "ISDIRT", 1)
                  .SetFocus
                  If lRow < .RowCount Then
                     .SetCurCell lRow, 1
                  ElseIf lRow > 1 Then
                     .SetCurCell lRow - 1, 1
                  End If
                  For lRow = 1 To .RowCount - 1
                     .CellValue(lRow, 1) = lRow
                  Next
               End If
            End With
         End If
         Call CarregaFechamento
         
      End With
      Set TlVenda = Nothing
   End If
   
Saida:
   mvarMe.CmdCancel.Enabled = True
   mvarMe.CmdOk.Enabled = True
   mvarMe.CmdSalvar.Enabled = True
   mvarMe.CmdVenda.Enabled = True
   If Not mvarCalEvent Is Nothing Then
      mvarCalEvent.CustomProperties("EXIBIRVENDA") = False
   End If
End Sub
Private Function VerificarPacote(pCliente As Boolean, ByRef pSaldo As Integer, Optional pVincular = False) As Boolean
   Dim Sql        As String
   Dim RsSaldo    As Object
   Dim sMsg       As String
   Dim sTit       As String
   Dim bVincular  As Boolean
   Dim nIDVENDA   As Integer
   Dim sDTVENDA   As String
   Dim nVLITENS   As Currency
   Dim nVLPGTO    As Currency
   Dim nQtd       As Integer
'   Dim nAux       As Integer
'   Dim MyAdd      As Object
'   Dim MySel      As Collection
   Dim SqlSld     As String
'   Dim dDTFIM      As Date
'
   Call SalvarAtendimento
   
   nIDVENDA = 0
   
   Dim oTLVinc As TL_VincularVenda
   
   '******
   '* Verifica se o atendimento possui serviço sem baixa
   SqlSld = "Select Count(*) [SALDO]" & vbNewLine
   SqlSld = SqlSld & " From OSESSAO" & vbNewLine
   SqlSld = SqlSld & " Where IDLOJA=" & mvarIDLOJA & vbNewLine
   SqlSld = SqlSld & " And IDATENDIMENTO=" & mvarIDATENDIMENTO & vbNewLine
   SqlSld = SqlSld & " And IDTPSERVICO<>1" & vbNewLine
   SqlSld = SqlSld & " And (isNull(IDVENDA,0)=0 Or isNull(IDITEM,0)=0)"
   If mvarSys.xDb.AbreTabela(SqlSld, RsSaldo) Then
      pSaldo = RsSaldo("SALDO")
   End If
   If pSaldo > 0 Then
      Set oTLVinc = New TL_VincularVenda
      Set oTLVinc.Sys = mvarSys
      oTLVinc.IDLOJA = mvarIDLOJA
      oTLVinc.IDATENDIMENTO = mvarIDATENDIMENTO
      
      Sql = oTLVinc.GetSqlSaldoVenda(pCliente, True)
      
      If mvarSys.xDb.AbreTabela(Sql, RsSaldo) Then
         nQtd = RsSaldo.RecordCount
         nIDVENDA = xVal(RsSaldo("IDVENDA") & "")
         sDTVENDA = RsSaldo("DTVENDA") & ""
         nVLITENS = xVal(RsSaldo("VLITENS") & "")
         nVLPGTO = xVal(RsSaldo("VLPGTO") & "")
         
         pSaldo = 0
         While Not RsSaldo.EOF
            pSaldo = pSaldo + xVal(RsSaldo("SALDO") & "")
            RsSaldo.MoveNext
         Wend
         If pSaldo > 0 Then
            If Not pVincular Then
               VerificarPacote = True
               
               sMsg = "Cliente possui saldo de serviço (" & pSaldo & ")"
               sMsg = sMsg & " em vendas anteriores." & vbNewLine & vbNewLine
               sMsg = sMsg & "Deseja vincular agora?"
               sTit = "ATENÇÃO!!   (Saldo: " & pSaldo & ")"
               
               bVincular = (vbYes = ExibirPergunta(sMsg, sTit))
            Else
               bVincular = True
            End If
         End If
      End If
   End If
   
   pSaldo = 0
   If bVincular Then
      Call mvarMe_CmdVincularClick
'      '*************
'      '* Selecionar Vendas
'      If nQtd > 1 Then
'
'         Set MyAdd = CriarObjeto("XActive.DsAdd", False)
'         If Not MyAdd Is Nothing Then
'            With MyAdd
'               .Limpar
'               Set .xDb = mvarSys.xDb
'               .Caption = "Vendas com Saldo."
'               .Chaves.Add "IDVENDA", "IDVENDA"
'               .Chaves.Add "DTVENDA", "DTVENDA"
'               .Chaves.Add "VENDA", "VENDA"
'               .Chaves.Add "PGTO", "PGTO"
'
'               With .TITULO
'                  '0-eAlin.Esquerda 1-eAlin.Direita 2-eAlin.Centro)
'                  Call .Add("IDVENDA", "Nº Venda", 8, 2)
'                  Call .Add("DTVENDA", "Data", 10, 2)
'                  Call .Add("NOME", "Cliente", 20, 0)
'                  Call .Add("PRODUTO", "Produto", 15, 2)
'                  Call .Add("VENDA", "Valor", 10, 1)
'                  Call .Add("PGTO", "Pagamento", 10, 1)
'                  Call .Add("SALDO", "Saldo", 8, 1)
'               End With
'               .CheckBox = True
'               .Query = Sql
'
'               .Show
'               Set MySel = .Selecao
'               If MySel Is Nothing Then
'                  nIDVENDA = 0
'                  pSaldo = 0
'               Else
'                  If MySel.Count = 0 Then
'                     nIDVENDA = 0
'                     pSaldo = 0
'                  Else
'                     nIDVENDA = MySel(1)("IDVENDA")
'                     sDTVENDA = MySel(1)("DTVENDA")
'                     nVLVENDA = MySel(1)("VENDA")
'                     nVLPGTO = MySel(1)("PGTO")
'                     'nAux = xVal(RsSaldo("IDVENDA") & "")
'                     'While Not RsSaldo.EOF Or nIDVENDA <> nAux
'                     '   nAux = xVal(RsSaldo("IDVENDA") & "")
'                     '   RsSaldo.MoveNext
'                     'Wend
'                     'If nIDVENDA <> nAux Then nIDVENDA = 0
'                  End If
'               End If
'            End With
'         End If
'      ElseIf nQtd = 1 Then
'         nIDVENDA = nIDVENDA
'      End If
'
'      If nIDVENDA <> 0 Then
'         nAux = pSaldo
'         Call VincularVenda(nIDVENDA, pSaldo, sDTVENDA, nVLVENDA, nVLPGTO)
'         'VerificarPacote = (nAux <> pSaldo)
'      End If
   End If
   If mvarSys.xDb.AbreTabela(SqlSld, RsSaldo) Then
      pSaldo = RsSaldo("SALDO")
   End If
End Function
'Private Sub VincularVenda(nIDVENDA As Integer, ByRef pSaldo As Integer, pDTVENDA As String, pVlItens As Currency, pVLPGTO As Currency, Optional nIDMOEDA As Integer)
'   Dim Sql     As String
'   Dim MyAdd   As Object
'   Dim MySel   As Collection
'   Dim TbMoeda As Object
'   Dim i       As Integer
'   Dim sTag    As String
'   Dim n       As Variant
'   Dim sMsg    As String
'   Dim nResult As Integer
'
'   Screen.MousePointer = vbHourglass
'
'   Sql = "Select Distinct I.IDSESSAO, S.DSCSERVICO [Serviço], T.DSCTRATAMENTO [Tratamento], A.DSCAREA [Área]" & vbNewLine
'   Sql = Sql & ", I.IDVENDA [Venda]" & vbNewLine
'   Sql = Sql & ", I.IDTPSERVICO, I.IDTPTRATAMENTO, I.IDAREA" & vbNewLine
'   Sql = Sql & ", (" & SqlSaldo(CStr(nIDVENDA)) & ") [Saldo]" & vbNewLine
'   Sql = Sql & " From OSESSAO I " & vbNewLine
'   Sql = Sql & " Join OTPSERVICO S     On I.IDLOJA=S.IDLOJA And I.IDTPSERVICO=S.IDTPSERVICO" & vbNewLine
'   Sql = Sql & " Join OTPTRATAMENTO T  On I.IDLOJA=T.IDLOJA And I.IDTPTRATAMENTO=T.IDTPTRATAMENTO" & vbNewLine
'   Sql = Sql & " Join OTRATAMENTO_PROD TP on TP.IDLOJA=I.IDLOJA And (TP.IDTPTRATAMENTO=I.IDTPTRATAMENTO or TP.IDTPTRATAMENTO=0)" & vbNewLine
'   Sql = Sql & " Join SPRODUTO P On TP.IDLOJA=P.IDLOJA And TP.IDPROD=P.IDPROD AND P.VLVENDA>0" & vbNewLine
'   Sql = Sql & " Left Join OAREA A          On I.IDLOJA=A.IDLOJA And I.IDAREA=A.IDAREA" & vbNewLine
'   Sql = Sql & " Where I.IDLOJA = " & mvarIDLOJA & vbNewLine
'   If mvarIDATENDIMENTO <> 0 Then
'      Sql = Sql & " And I.IDATENDIMENTO= " & mvarIDATENDIMENTO
'   End If
'   If Not mvarSys.xDb.AbreTabela(Sql) Then
'      Dim Sql2  As String
'      Dim MyRS As Object
'      Dim MyRs2 As Object
'      Dim nIdProd As Long
'      Dim sNMProd As String
'      Dim nIDTPTRATAMENTO As Long
'
'      Sql2 = "Select Distinct I.IDLOJA, I.IDTPSERVICO, I.IDTPTRATAMENTO , I.IDAREA"
'      Sql2 = Sql2 & ", T.DSCTRATAMENTO"
'      Sql2 = Sql2 & " From OSESSAO I " & vbNewLine
'      Sql2 = Sql2 & " Join OTPSERVICO S     On I.IDLOJA=S.IDLOJA And I.IDTPSERVICO=S.IDTPSERVICO" & vbNewLine
'      Sql2 = Sql2 & " Join OTPTRATAMENTO T  On I.IDLOJA=T.IDLOJA And I.IDTPTRATAMENTO=T.IDTPTRATAMENTO" & vbNewLine
'      Sql2 = Sql2 & " Left Join OTRATAMENTO_PROD TP On I.IDLOJA=TP.IDLOJA And I.IDTPTRATAMENTO=TP.IDTPTRATAMENTO" & vbNewLine
'      Sql2 = Sql2 & " Left Join SPRODUTO P On TP.IDLOJA=P.IDLOJA And TP.IDPROD=P.IDPROD AND P.VLVENDA>0" & vbNewLine
'      Sql2 = Sql2 & " Left Join OAREA A          On I.IDLOJA=A.IDLOJA And I.IDAREA=A.IDAREA" & vbNewLine
'      Sql2 = Sql2 & " Where I.IDLOJA = " & mvarIDLOJA & vbNewLine
'      Sql2 = Sql2 & " And I.IDATENDIMENTO= " & mvarIDATENDIMENTO
'      If mvarSys.xDb.AbreTabela(Sql2, MyRS) Then
'         While Not MyRS.EOF
'            nIdProd = 0
'            sNMProd = MyRS("DSCTRATAMENTO")
'            nIDTPTRATAMENTO = MyRS("IDTPTRATAMENTO")
'            Sql2 = "Select IDPROD"
'            Sql2 = Sql2 & " From SPRODUTO"
'            Sql2 = Sql2 & " Where NMPROD Like '" & sNMProd & "'"
'            If mvarSys.xDb.AbreTabela(Sql2, MyRs2) Then
'               nIdProd = MyRs2("IDPROD")
'            End If
'            Call SalvarProd(nIdProd, sNMProd, nIDTPTRATAMENTO)
'            MyRS.MoveNext
'         Wend
'      End If
'   End If
'
'   Set MyAdd = CriarObjeto("XActive.DsAdd", False)
'   If Not MyAdd Is Nothing Then
'      With MyAdd
'         .Limpar
'         Set .xDb = mvarSys.xDb
'         .Caption = "Itens do Atendimento."
'         .Chaves.Add "IDSESSAO", "IDSESSAO"
'
'         With .TITULO
'            Call .Add("IDSESSAO", "Item", 8, 2) 'eAlin.Centro)
'            Call .Add("Serviço", "Serviço", 10, 0) 'eAlin.Esquerda)
'            Call .Add("Tratamento", "Tratamento", IIf(mvarSys.Propriedades("TPAREA") = "0", 40, 20), 0) 'eAlin.Esquerda)
'            Call .Add("Área", "Área", IIf(mvarSys.Propriedades("TPAREA") = "0", 0, 20), 0) 'eAlin.Esquerda)
'            Call .Add("Saldo", "Saldo", 10, 0) 'eAlin.Esquerda)
'            'Call .Add("IDTPSERVICO", "IDTPSERVICO", 10, 1) 'eAlin.Direita)
'            'Call .Add("IDTPTRATAMENTO", "IDTPTRATAMENTO", 10, 1) 'eAlin.Direita)
'            'Call .Add("IDAREA", "IDAREA", 10, 2) 'eAlin.Centro)
'
'         End With
'         .CheckBox = True
'         .Query = Sql
'
'         .Show
'         Set MySel = .Selecao
'
'      End With
'      If MySel Is Nothing Then
'         nResult = pSaldo
'      Else
'         nResult = pSaldo - MySel.Count
'      End If
'   End If
'   If nResult < 0 Then
'      sMsg = "Seleção Inválida." & vbNewLine & vbNewLine
'      sMsg = sMsg & "Só poderá ser selecionada '" & pSaldo & "' itens."
'      Call ExibirAviso(sMsg, "Atenção!      (Saldo: '" & pSaldo & "')")
'      GoTo Saida
'   End If
'   '***********
'   '* Vincular Venda
'   If Not MySel Is Nothing Then
'      If MySel.Count > 0 Then
'         If nIDMOEDA = 0 Then
'            nIDMOEDA = xVal(mvarSys.Propriedades("IDMOEDA"))
'         End If
'         nIDMOEDA = IIf(nIDMOEDA = 0, 1, nIDMOEDA)
'         If mvarSys.Propriedades("IDMOEDA") <> nIDMOEDA Or nIDMOEDA = 0 Then
'            Set TbMoeda = CriarObjeto("BANCO_3R.TB_FMOEDA")
'            Set TbMoeda.xDb = mvarSys.xDb
'            Call TbMoeda.Pesquisar(Ch_IDMOEDA:=nIDMOEDA)
'            mvarSys.Propriedades("IDMOEDA") = TbMoeda.IDMOEDA
'            mvarSys.Propriedades("SIGLAMOEDA") = TbMoeda.SIGLAMOEDA
'            Set TbMoeda = Nothing
'         End If
'
'         Call IncluiGrdLinhaVenda
'         i = mvarMe.GrdVenda.RowCount - 1
'         If Not IsMissing(mvarMe.GrdVenda.RowTag(i)) Then sTag = mvarMe.GrdVenda.RowTag(i)
'         Call SetTag(sTag, "ITEM", i)
'         Call SetTag(sTag, "IDVENDA", StrZero(nIDVENDA, mvarSys.Propriedades("LENIDVENDA")))
'         Call SetTag(sTag, "DTVENDA", pDTVENDA)
'         Call SetTag(sTag, "VLITENS", mvarSys.Propriedades("SIGLAMOEDA") & " " & ValBr(pVlItens))
'         Call SetTag(sTag, "VLPGTO", mvarSys.Propriedades("SIGLAMOEDA") & " " & ValBr(pVLPGTO))
'         Call SetTag(sTag, "SALDO", nResult)
'
'         mvarMe.GrdVenda.RowTag(i) = sTag
'
'         mvarMe.GrdVenda.CellValue(i, "ITEM") = i
'         mvarMe.GrdVenda.CellValue(i, "IDVENDA") = StrZero(nIDVENDA, mvarSys.Propriedades("LENIDVENDA"))
'         mvarMe.GrdVenda.CellValue(i, "DTVENDA") = pDTVENDA
'         mvarMe.GrdVenda.CellValue(i, "VLITENS") = mvarSys.Propriedades("SIGLAMOEDA") & " " & ValBr(pVlItens)
'         mvarMe.GrdVenda.CellValue(i, "VLPGTO") = mvarSys.Propriedades("SIGLAMOEDA") & " " & ValBr(pVLPGTO)
'
'         Call SetTag(mvarMe.GrdVenda, "ISDIRT", 1)
'         Call SalvarAtendimento
'      End If
'
'      '***********
'      '* Vincular Itens de Atendimento
'      For Each n In MySel
'         Sql = "Update OSESSAO "
'         Sql = Sql & " Set IDVENDA= " & SqlNum(nIDVENDA)
'         Sql = Sql & ", ALTERSTAMP=1"
'         Sql = Sql & ", TIMESTAMP = GetDate()"
'         Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
'         Sql = Sql & " And IDATENDIMENTO=" & SqlNum(mvarIDATENDIMENTO)
'         Sql = Sql & " And IDSESSAO=" & SqlNum(n("IDSESSAO"))
'         Call mvarSys.xDb.Executa(Sql)
'      Next
'   End If
'   pSaldo = nResult
'Saida:
'   Screen.MousePointer = vbDefault
'End Sub
Private Function SalvarProd(pIdProd As Long, pNmProd As String, pIDTPTRATAMENTO As Long, Optional pVlVenda As Currency) As Boolean
   Dim TbProd As Object
   Dim TbTratProd As Object
   Dim bExiste As Boolean
      
   If Trim(pNmProd) <> "" Then
      Set TbProd = Nothing
      Set TbProd = CriarObjeto("BANCO_3R.TB_SPRODUTO", False)
      Set TbProd.xDb = mvarSys.xDb
      bExiste = TbProd.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPROD:=pIdProd)
      TbProd.IDLOJA = mvarIDLOJA
      TbProd.CODPROD = pIdProd
      TbProd.NMPROD = pNmProd
      If pVlVenda <> 0 Then
         TbProd.VLVENDA = pVlVenda
      End If
      If TbProd.DSCPROD = "" Then
         TbProd.DSCPROD = TbProd.NMPROD
      End If
      If Not bExiste Then
         TbProd.DTCADASTRO = mvarSys.xDb.Sysdate()
         TbProd.UNIDCONTROLE = "un"
         TbProd.UNIDCOMPRA = TbProd.UNIDCONTROLE
         TbProd.UNIDVENDA = TbProd.UNIDCONTROLE
         TbProd.VLMEDIO = TbProd.VLVENDA
         TbProd.IDUSU = mvarSys.IDUSU
      End If
      TbProd.EVENDA = 1
      TbProd.ESERVICO = 1
      TbProd.SITPROD = "A"
      If TbProd.Salvar Then
         Set TbTratProd = Nothing
         Set TbTratProd = CriarObjeto("BANCO_3R.TB_OTRATAMENTO_PROD", False)
         Set TbTratProd.xDb = mvarSys.xDb
         Call TbTratProd.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDTPTRATAMENTO=" & pIDTPTRATAMENTO)
         TbTratProd.IDLOJA = mvarIDLOJA
         TbTratProd.IDPROD = TbProd.IDPROD
         TbTratProd.IDTPSERVICO = 2
         TbTratProd.IDTPTRATAMENTO = pIDTPTRATAMENTO
         TbTratProd.IDAREA = 0
         
         Call TbTratProd.Salvar
      End If
   End If
End Function
Private Function SqlSaldo(pIDVENDA As String) As String
   Dim Sql As String
   
   Sql = "Select Distinct Cast(( IsNull((SELECT SUM(I.QTDVENDA)" & vbNewLine
   Sql = Sql & " FROM CITENSVENDA I" & vbNewLine
   Sql = Sql & " JOIN SPRODUTO P ON I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD And P.ESERVICO=1 And P.EVENDA=1" & vbNewLine
   Sql = Sql & " Where I.IDLOJA = V.IDLOJA" & vbNewLine
   Sql = Sql & " AND I.IDVENDA=V.IDVENDA),0)" & vbNewLine
   Sql = Sql & " - " & vbNewLine
   Sql = Sql & " IsNull((SELECT COUNT(S.IDSESSAO)" & vbNewLine
   Sql = Sql & " From OSESSAO S" & vbNewLine
   Sql = Sql & " Join OATENDIMENTO_VENDA AV On AV.IDLOJA=S.IDLOJA" & vbNewLine
   Sql = Sql & "  And AV.IDATENDIMENTO=S.IDATENDIMENTO" & vbNewLine
   Sql = Sql & " And AV.IDVENDA=S.IDVENDA" & vbNewLine
   Sql = Sql & " And S.IDTPSERVICO<>1" & vbNewLine
   Sql = Sql & " Where S.IDLOJA = V.IDLOJA" & vbNewLine
   Sql = Sql & " AND S.IDVENDA=V.IDVENDA),0)" & vbNewLine
   Sql = Sql & ") as Integer) [Sld]" & vbNewLine
   Sql = Sql & " From CVENDA V" & vbNewLine
   Sql = Sql & " Join OCLIENTE C On C.IDLOJA=V.IDLOJA And V.IDCLIENTE=C.IDCLIENTE" & vbNewLine
   Sql = Sql & " Left Join FRECIBO R On V.IDLOJA=R.IDLOJA And V.IDVENDA=R.IDVENDA" & vbNewLine
   Sql = Sql & " Left Join FNOTAFISCAL N On V.IDLOJA=N.IDLOJA And V.IDVENDA=N.IDVENDA And N.FLGCANCELADA=0" & vbNewLine
   Sql = Sql & " Where V.IDLOJA = " & mvarIDLOJA & vbNewLine
   Sql = Sql & " And V.IDVENDA IN (" & pIDVENDA & " )" & vbNewLine
   Sql = Sql & " Union" & vbNewLine
   Sql = Sql & " Select Distinct Cast(( IsNull((SELECT SUM(I.QTDVENDA)" & vbNewLine
   Sql = Sql & " FROM CITENSVENDA I" & vbNewLine
   Sql = Sql & " JOIN SPRODUTO P ON I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD And P.ESERVICO=1 And P.EVENDA=1" & vbNewLine
   Sql = Sql & " Where i.IDLOJA = V.IDLOJA" & vbNewLine
   Sql = Sql & " AND I.IDVENDA=V.IDVENDA),0)" & vbNewLine
   Sql = Sql & " - " & vbNewLine
   Sql = Sql & " IsNull((SELECT COUNT(S.IDSESSAO)" & vbNewLine
   Sql = Sql & " From OSESSAO S" & vbNewLine
   Sql = Sql & " Join OATENDIMENTO_VENDA AV On AV.IDLOJA=S.IDLOJA" & vbNewLine
   Sql = Sql & "  And AV.IDATENDIMENTO=S.IDATENDIMENTO" & vbNewLine
   Sql = Sql & "  And AV.IDVENDA=S.IDVENDA" & vbNewLine
   Sql = Sql & "  And S.IDTPSERVICO<>1" & vbNewLine
   Sql = Sql & " Where S.IDLOJA = V.IDLOJA" & vbNewLine
   Sql = Sql & " AND S.IDVENDA=V.IDVENDA),0)" & vbNewLine
   Sql = Sql & ") as Integer)" & vbNewLine
   Sql = Sql & " From CVENDA V" & vbNewLine
   Sql = Sql & " Join CITENSVENDA I on V.IDLOJA=I.IDLOJA and V.IDVENDA=I.IDVENDA" & vbNewLine
   Sql = Sql & " Join OATENDIMENTO_VENDA AV on AV.IDLOJA=V.IDLOJA and AV.IDVENDA=V.IDVENDA" & vbNewLine
   Sql = Sql & " Join OATENDIMENTO A on AV.IDLOJA=A.IDLOJA and AV.IDATENDIMENTO=A.IDATENDIMENTO" & vbNewLine
   Sql = Sql & " Join OCLIENTE C on A.IDLOJA=C.IDLOJA and A.IDCLIENTE=C.IDCLIENTE" & vbNewLine
   Sql = Sql & " Left Join FRECIBO R On V.IDLOJA=R.IDLOJA And V.IDVENDA=R.IDVENDA" & vbNewLine
   Sql = Sql & " Left Join FNOTAFISCAL N On V.IDLOJA=N.IDLOJA And V.IDVENDA=N.IDVENDA And N.FLGCANCELADA=0" & vbNewLine
   Sql = Sql & " Where V.IDLOJA = " & mvarIDLOJA & vbNewLine
   Sql = Sql & " And V.IDVENDA IN (" & pIDVENDA & ")" & vbNewLine
   
   SqlSaldo = Sql
End Function
Private Sub ExibirFatura()
   Dim TlFatura As Object
   Dim NgFinanc As Object
   Dim AtendSalvo As Boolean
  
   On Error GoTo Saida
        
   If Not VerificaAtendimento(bVerifVenda:=False, bParaVenda:=True) Then Exit Sub
   
   
   AtendSalvo = SalvarAtendimento
   
   If AtendSalvo Then
      Set TlFatura = CriarObjeto("Financ3R.TL_Fatura")
      With TlFatura
         Set .Sys = mvarSys
         'Set .CalControl = mvarCalControl
         
         .IDLOJA = mvarIDLOJA
         .IDATENDIMENTO = mvarIDATENDIMENTO
         .IDCLIENTE = mvarIDCLIENTE
         
         
         .Show vbModal
         
         Call CarregaFechamento
         
      End With
      Set TlFatura = Nothing
   End If
   
Saida:
   mvarMe.CmdCancel.Enabled = True
   mvarMe.CmdOk.Enabled = True
   mvarMe.CmdSalvar.Enabled = True
   mvarMe.CmdVenda.Enabled = True
End Sub
Private Sub mvarMe_CmdVendaDropDown()
   Dim xMenu      As CommandBar
   Dim oMenuItem  As CommandBarControl

   '**************
   '* Pop-up Edit
   On Error Resume Next
   If ContextMenu Is Nothing Then
      Set ContextMenu = mvarSys.MDI.CommandBars
   End If
   mvarSys.MDI.CommandBars.ContextMenus(99).Delete
   
   Set xMenu = ContextMenu.ContextMenus.Add(99, "Menu Aux")
   With xMenu
      Set oMenuItem = AddButtonBar(.Controls, 1, "Venda", Category:="ATENDIMENTO")
      'oMenuItem.Controls.Add XTPControlType.xtpControlButtonPopup, 11, "teste"
      AddButtonBar .Controls, 2, "Fatura", True, Category:="ATENDIMENTO"
   End With

   With mvarMe.CmdVenda
      xMenu.ShowPopup , .Container.Left + .Left, .Container.Top + .Top + .Height + .Height + 40
      'mvarMe.PopupMenu ContextMenu, , 0, 0
      'mvarMe.PopupMenu ContextMenu, , .Container.Left + .Left, .Container.Top + .Top + .Height
   End With
End Sub

Private Sub mvarMe_CmdVincularClick()
   Dim oTLVinc As TL_VincularVenda
      
   
   
   Dim n       As Object
   Dim MyAdd   As Object 'DsAdd
   Dim MySel   As Collection
   Dim Sql     As String
    
   If mvarIDATENDIMENTO = 0 Then
      If Not VerificaAtendimento(bVerifVenda:=mvarMe.CmdChave.Visible, bExibeMsg:=Not bChave) Or bChave Then
         Exit Sub
      Else
         DoEvents
         If Not SalvarAtendimento Then
            Exit Sub
         End If
      End If
   End If
    
'   Call SalvarVinculoVenda
   
   'If xVal(mvarSys.Propriedades("AJUSTAVENDA")) = 0 Then
   '   If AjustaAtendimentoVenda Then
   '      mvarSys.Propriedades("AJUSTAVENDA") = 1
   '   Else
   '      'While Not AjustaAtendimentoVenda
   '      'Wend
   '   End If
   'End If

   Set oTLVinc = New TL_VincularVenda
   Set oTLVinc.Sys = mvarSys
   oTLVinc.IDLOJA = mvarIDLOJA
   oTLVinc.IDATENDIMENTO = mvarIDATENDIMENTO
   oTLVinc.Show
   
   Call PopulaGRDVENDA
End Sub
Private Function AjustaAtendimentoVenda() As Boolean
   Dim Sql  As String
   Dim nVs  As Integer
   Dim bOk  As Boolean
   Dim MyRS As Object
   Dim MyRs2 As Object
   Dim MyRs3 As Object
   Dim nQtdVenda As Integer
   Dim nQtdAux As Integer
   Dim nIDATEND As Integer
   
   AjustaAtendimentoVenda = True
   nVs = 0
   With mvarSys.xDb
      Sql = "Select ATUBD From VERSAOBD"
      If .AbreTabela(Sql, MyRS) Then
         nVs = MyRS("ATUBD")
      End If
      If nVs <> 0 Then 'And nVs < 55 Then
         Sql = "UPDATE OSESSAO" & vbNewLine
         Sql = Sql & " SET IDVENDA=0 " & vbNewLine
         Sql = Sql & ", ALTERSTAMP=1" & vbNewLine
         Sql = Sql & ", TIMESTAMP=GetDate()" & vbNewLine
         Sql = Sql & " WHERE IDVENDA IS NULL"
         bOk = .Executa(Sql)
         
         Sql = "UPDATE OSESSAO " & vbNewLine
         Sql = Sql & " SET OSESSAO.IDVENDA =(SELECT AV.IDVENDA FROM OATENDIMENTO_VENDA AV WHERE OSESSAO.IDLOJA=AV.IDLOJA AND OSESSAO.IDATENDIMENTO=AV.IDATENDIMENTO)" & vbNewLine
         Sql = Sql & "Where OSESSAO.IDLOJA = " & mvarIDLOJA & vbNewLine
         Sql = Sql & " AND OSESSAO.IDATENDIMENTO IN" & vbNewLine
         Sql = Sql & "    (SELECT AV2.IDATENDIMENTO " & vbNewLine
         Sql = Sql & "     FROM OATENDIMENTO_VENDA AV2 " & vbNewLine
         Sql = Sql & "     WHERE AV2.IDLOJA =" & mvarIDLOJA & vbNewLine
         Sql = Sql & "     And AV2.IDATENDIMENTO<>0" & mvarIDLOJA & vbNewLine
         Sql = Sql & "     GROUP BY AV2.IDATENDIMENTO " & vbNewLine
         Sql = Sql & "     HAVING COUNT(AV2.IDVENDA)=1)" & vbNewLine
         Sql = Sql & " AND OSESSAO.IDVENDA =0"
         Sql = Sql & " AND OSESSAO.IDTPSERVICO<>1"
         bOk = bOk And .Executa(Sql)
         
'         Sql = " Select Distinct Right(REPLICATE('0',6)+ Cast(V.IDVENDA as varchar),6) [IDVENDA]"
'         Sql = Sql & " , Convert(varchar, V.DTVENDA, 103) [DTVENDA], C.TEL1, C.NOME"
'         Sql = Sql & " , V.VLVENDA, V.VLPGTO, V.IDMOEDA, Sum(Distinct I.QTDVENDA)-(Select Count(S.IDSESSAO) From OSESSAO S where S.IDLOJA=V.IDLOJA and S.IDVENDA=V.IDVENDA) [SALDO]"
'         Sql = Sql & " From CVENDA V"
'         Sql = Sql & " Join CITENSVENDA I on V.IDLOJA=I.IDLOJA and V.IDVENDA=I.IDVENDA"
'         Sql = Sql & " Join SPRODUTO P on I.IDLOJA=P.IDLOJA and I.IDPROD=P.IDPROD"
'         Sql = Sql & " Join OCLIENTE C on V.IDLOJA=C.IDLOJA and C.IDCLIENTE=V.IDCLIENTE"
'         Sql = Sql & " Where V.IDLOJA =  " & mvarIDLOJA
'         Sql = Sql & " And P.ESERVICO = 1"
'         Sql = Sql & " Group by V.IDLOJA, V.DTVENDA, V.IDVENDA,C.TEL1, C.NOME, V.VLVENDA, V.VLPGTO, V.IDMOEDA"
'         Sql = Sql & " HavIng Sum(Distinct I.QTDVENDA) >"
'         Sql = Sql & " (Select Count(S.IDSESSAO) from OSESSAO s where S.IDLOJA=V.IDLOJA and S.IDVENDA=V.IDVENDA)"
'         Sql = Sql & " Order By 1"
         '*********
         '* Limpa Sessão com vendas não associadas
         Sql = ""
         Sql = Sql & "Select DISTINCT A.IDLOJA, A.IDATENDIMENTO, S.IDSESSAO"
         Sql = Sql & " From OATENDIMENTO A"
         Sql = Sql & " JOIN OSESSAO S On A.IDLOJA=S.IDLOJA And A.IDATENDIMENTO=S.IDATENDIMENTO"
         Sql = Sql & " Where A.IDLOJA =" & mvarIDLOJA
         Sql = Sql & " And A.IDATENDIMENTO<>0"
         Sql = Sql & " And S.IDTPSERVICO<>1"
         Sql = Sql & " AND S.IDVENDA NOT IN (SELECT AV.IDVENDA FROM OATENDIMENTO_VENDA AV WHERE  AV.IDLOJA=A.IDLOJA And AV.IDATENDIMENTO=A.IDATENDIMENTO)"
         If .AbreTabela(Sql, MyRS) Then
            While Not MyRS.EOF
               Sql = "UPDATE OSESSAO" & vbNewLine
               Sql = Sql & " SET IDVENDA=0 " & vbNewLine
               Sql = Sql & ", ALTERSTAMP=1" & vbNewLine
               Sql = Sql & ", TIMESTAMP=GetDate()" & vbNewLine
               Sql = Sql & " WHERE IDLOJA=" & MyRS("IDLOJA")
               Sql = Sql & " AND IDATENDIMENTO=" & MyRS("IDATENDIMENTO")
               Sql = Sql & " AND IDSESSAO=" & MyRS("IDSESSAO")
               bOk = bOk And .Executa(Sql)
               
               MyRS.MoveNext
            Wend
         End If
                 
         Sql = "Select A.IDLOJA, A.IDATENDIMENTO, A.DTATEND, A.IDCLIENTE, S.IDSESSAO" & vbNewLine
         Sql = Sql & " , S.IDVENDA, S.IDTPSERVICO, S.IDTPTRATAMENTO, S.IDAREA" & vbNewLine
         Sql = Sql & " From OATENDIMENTO A" & vbNewLine
         Sql = Sql & " JOIN OSESSAO S On A.IDLOJA=S.IDLOJA And A.IDATENDIMENTO=S.IDATENDIMENTO" & vbNewLine
         Sql = Sql & " Where A.IDLOJA = " & mvarIDLOJA & vbNewLine
         Sql = Sql & " And A.IDATENDIMENTO<>0"
         Sql = Sql & " And S.IDVENDA =0" & vbNewLine
         Sql = Sql & " And S.IDTPSERVICO<>1" & vbNewLine
         Sql = Sql & " ORDER BY A.IDLOJA, A.IDATENDIMENTO, A.DTATEND, S.IDSESSAO" & vbNewLine
         If .AbreTabela(Sql, MyRS) Then
            While Not MyRS.EOF
               '**************
               '* Atualizando Itens com vendas associadas
               Sql = ""
               Sql = Sql & " Select V.IDLOJA, V.IDVENDA, Sum(I.QTDVENDA) [QTDVENDA]" & vbNewLine
               Sql = Sql & " From CVENDA V" & vbNewLine
               Sql = Sql & " JOIN CITENSVENDA I On V.IDLOJA=I.IDLOJA And V.IDVENDA=I.IDVENDA" & vbNewLine
               Sql = Sql & " JOIN SPRODUTO P On I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD AND P.EVENDA=1 AND P.ESERVICO=1" & vbNewLine
               Sql = Sql & " JOIN OATENDIMENTO_VENDA AV On V.IDLOJA=AV.IDLOJA And V.IDVENDA=AV.IDVENDA" & vbNewLine
               Sql = Sql & " Where AV.IDLOJA = " & MyRS("IDLOJA") & vbNewLine
               Sql = Sql & " And AV.IDATENDIMENTO =" & MyRS("IDATENDIMENTO") & vbNewLine
               Sql = Sql & " GROUP BY V.IDLOJA, V.IDVENDA" & vbNewLine
               Sql = Sql & " ORDER BY V.IDLOJA, V.IDVENDA" & vbNewLine
               If .AbreTabela(Sql, MyRs2) Then
                  While Not MyRs2.EOF And Not MyRS.EOF
                     Sql = ""
                     Sql = Sql & " SELECT COUNT(*) [QTD]"
                     Sql = Sql & " FROM OSESSAO "
                     Sql = Sql & " WHERE IDLOJA =" & MyRS("IDLOJA")
                     Sql = Sql & " AND IDVENDA=" & MyRs2("IDVENDA")
                     Sql = Sql & " AND IDTPSERVICO<>1"
                     nQtdVenda = 0
                     If .AbreTabela(Sql, MyRs3) Then
                        nQtdVenda = MyRs3("QTD")
                     End If
                     
                     nQtdAux = MyRs2("QTDVENDA") - nQtdVenda
                     nIDATEND = MyRS("IDATENDIMENTO")
                     While nQtdAux > 0
                        If nQtdAux <> MyRs2("QTDVENDA") Then
                           If nIDATEND <> MyRS("IDATENDIMENTO") Then
                              nQtdAux = 0
                           End If
                        End If
                        If nQtdAux > 0 Then
                           Sql = "Update OSESSAO"
                           Sql = Sql & " Set IDVENDA=" & MyRs2("IDVENDA")
                           Sql = Sql & ", ALTERSTAMP=1"
                           Sql = Sql & ", TIMESTAMP=GetDate()"
                           Sql = Sql & " Where IDLOJA=" & MyRS("IDLOJA")
                           Sql = Sql & " And IDATENDIMENTO=" & MyRS("IDATENDIMENTO")
                           Sql = Sql & " And IDSESSAO=" & MyRS("IDSESSAO")
                           Sql = Sql & " AND IDTPSERVICO<>1"

                           bOk = bOk And .Executa(Sql)
                           If bOk Then
                              nQtdAux = nQtdAux - 1
                              nIDATEND = MyRS("IDATENDIMENTO")
                              MyRS.MoveNext
                           End If
                        End If
                     Wend
                     MyRs2.MoveNext
                  Wend
               End If
               '**************
               '* Procurar venda sem atendimento associado
               If mvarSys.Propriedades("TPAREA") = "1" Or mvarSys.Propriedades("TPAREA") = "" Then
                  Sql = ""
                  Sql = Sql & " Select V.IDLOJA, V.IDVENDA, V.IDCLIENTE, Sum(I.QTDVENDA) [QTDVENDA]" & vbNewLine
                  Sql = Sql & " From CVENDA V" & vbNewLine
                  Sql = Sql & " JOIN CITENSVENDA I On V.IDLOJA=I.IDLOJA And V.IDVENDA=I.IDVENDA" & vbNewLine
                  Sql = Sql & " JOIN SPRODUTO P On I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD AND P.EVENDA=1 AND P.ESERVICO=1" & vbNewLine
                  Sql = Sql & " LEFT JOIN OATENDIMENTO_VENDA AV On V.IDLOJA=AV.IDLOJA And V.IDVENDA=AV.IDVENDA" & vbNewLine
                  Sql = Sql & " Where V.IDLOJA = " & MyRS("IDLOJA") & vbNewLine
                  Sql = Sql & " And V.IDCLIENTE= " & MyRS("IDCLIENTE") & vbNewLine
                  Sql = Sql & " And AV.IDATENDIMENTO IS NULL" & vbNewLine
                  Sql = Sql & " GROUP BY V.IDLOJA, V.IDVENDA, V.IDCLIENTE"
                  If .AbreTabela(Sql, MyRs2) Then
                     Sql = "Insert into OATENDIMENTO_VENDA"
                     Sql = Sql & " (IDLOJA, IDATENDIMENTO, IDVENDA)"
                     Sql = Sql & " Values "
                     Sql = Sql & " (" & MyRS("IDLOJA")
                     Sql = Sql & ", " & MyRS("IDATENDIMENTO")
                     Sql = Sql & ", " & MyRs2("IDVENDA")
                     Sql = Sql & " )"
                     bOk = bOk And .Executa(Sql)
                     AjustaAtendimentoVenda = False
                     
                     While Not MyRs2.EOF And Not MyRS.EOF
                        nQtdAux = MyRs2("QTDVENDA")
                        nIDATEND = MyRS("IDATENDIMENTO")
                        While nQtdAux > 0
                           If nQtdAux <> MyRs2("QTDVENDA") Then
                              If MyRS.EOF Then
                                 nQtdAux = 0
                              Else
                                 If nIDATEND <> MyRS("IDATENDIMENTO") Then
                                    nQtdAux = 0
                                 End If
                              End If
                           End If
                           If nQtdAux > 0 Then
                              Sql = "Update OSESSAO"
                              Sql = Sql & " Set IDVENDA=" & MyRs2("IDVENDA")
                              Sql = Sql & " Where IDLOJA=" & MyRS("IDLOJA")
                              Sql = Sql & " And IDATENDIMENTO=" & MyRS("IDATENDIMENTO")
                              Sql = Sql & " And IDSESSAO=" & MyRS("IDSESSAO")
                              Sql = Sql & " AND IDTPSERVICO<>1"
                              
                              bOk = bOk Or .Executa(Sql)
                              If bOk Then
                                 nQtdAux = nQtdAux - 1
                                 nIDATEND = MyRS("IDATENDIMENTO")
                                 MyRS.MoveNext
                              End If
                           End If
                           
                        Wend
                        MyRs2.MoveNext
                     Wend
                  End If
               End If
               If Not MyRS.EOF Then
                  MyRS.MoveNext
               End If
            Wend
         End If
      End If
   End With
End Function
Private Sub mvarMe_GrdPRODUTOAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   Dim MyPROD  As Object 'TB_SPRODUTO
   
   Set MyPROD = CriarObjeto("BANCO_3R.TB_SPRODUTO")
   Set MyPROD.xDb = mvarSys.xDb
   If MyPROD.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPROD:=mvarMe.GrdPRODUTO.CellValue(lRow, "IDPROD")) Then
      mvarMe.GrdPRODUTO.CellValue(lRow, "UNIDCONTROLE") = MyPROD.UNIDCONTROLE
   End If
   Set MyPROD = Nothing
End Sub
Private Sub mvarMe_GrdPRODUTOBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   If lConvErr <> 0 Then
      ExibirAviso "Valor inválido."
      eResult = igEditResProceed
   End If
End Sub
Private Sub mvarMe_GrdPRODUTOColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   bDoDefault = False
End Sub
Private Sub mvarMe_GrdPRODUTOLostFocus()
   Dim sTag As String
   Dim i    As Integer
   Dim j    As Integer
   
   With mvarMe.GrdPRODUTO
      '.Tag = SetTag(.Tag, "ISDIRT", 0)
      If .RowCount <= 1 Then .Tag = SetTag(.Tag, "ISDIRT", 1)
      For i = 1 To .RowCount - 1
         If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
         For j = 1 To .ColCount
            If GetTag(sTag, .ColKey(j), 0) <> .CellValue(i, j) Then .Tag = SetTag(.Tag, "ISDIRT", 1)
            If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
         Next
         If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
      Next
   End With
End Sub

Private Sub mvarMe_GrdPRODUTOMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
   
   With mvarMe.GrdPRODUTO
      If lRow = .RowCount Then
         If bVendaFechada Then
            Call ExibirInformacao("Venda fechada.")
            Exit Sub
         End If
         
         If lRow > 1 Then
            If Val(.CellValue(lRow - 1, "IDPROD")) = 0 Then Exit Sub
            If Val(.CellValue(lRow - 1, "QTDCOMPRA")) = 0 Then Exit Sub
         End If
         Call IncluiGrdLinhaProd
         
         If Val(.CellValue(lRow, "IDPROD")) = 0 Then
            Call .SetCurCell(lRow, "IDPROD")
         Else
            Call .SetCurCell(lRow, "QTDCOMPRA")
         End If
      End If
   End With
End Sub
Private Sub mvarMe_GrdPRODUTORequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   With mvarMe.GrdPRODUTO
      bCancel = (GetTag(.ColTag(lCol), "EDIT", True) = False)
      If .CellFmtString(lRow, lCol) <> "" Then sText = .CellValue(lRow, lCol)
      eTextEditOpt = GetTag(.ColTag(lCol), "EDITOPT", 0)
   End With
End Sub
Private Sub mvarMe_GrdPRODUTOValidate(Cancel As Boolean)
   mvarMe.GrdPRODUTO.CommitEdit
End Sub
Private Sub mvarMe_GrdSESSAOAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   Dim i As Integer
   Static ColVisible As Integer
   With mvarMe.GrdSESSAO
      If lCol < .ColCount Then
         Call .SetCurCell(lRow, lCol + 1)
      Else
         Call .SetCurCell(lRow, lCol - 1)
      End If
      Call .SetCurCell(lRow, lCol)
      
      If InArray(.ColKey(lCol), Array("POTENCIA", "DISPAROS", "TEMPOSESSAO")) Then
         If Not IsNumeric(.CellValue(lRow, lCol)) Then
            Call ExibirAviso("Valor inválido.")
            .CellValue(lRow, lCol) = ""
            Call .SetCurCell(lRow, lCol)
         Else
            If ColVisible = 0 Then
               For i = 1 To .ColCount
                  If .ColVisible(i) Then
                     ColVisible = i
                  End If
               Next
            End If
            If lCol = .ColCount Or lCol = ColVisible Then
               If lRow < .RowCount - 1 Then
                  Call .SetCurCell(lRow + 1, .ColIndex("POTENCIA"))
               End If
            Else
               Call .SetCurCell(lRow, lCol + 1)
            End If
               
         End If
      End If
   End With
End Sub

Private Sub mvarMe_GrdSESSAOBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   Dim i          As Integer
   Dim Sql        As String
   Dim nNUMSESSAO As Long
   Dim nIDDIRECAO As Long
   Dim nIDAREA    As Long
   Dim nIDTPTRAT  As Long
   Dim nTPSERV    As Long
   Dim nFOTOTIPO  As Integer
   
   With mvarMe.GrdSESSAO
      If lConvErr <> 0 Then
         ExibirAviso "Valor inválido."
         eResult = igEditResProceed
      ElseIf .ColVisible("IDAREA") Then
         If .ColKey(lCol) = "IDAREA" Then
            nIDTPTRAT = .CellValue(lRow, "IDTPTRATAMENTO")
            For i = 1 To .RowCount - 2
               If vNewValue = .CellValue(i, lCol) And nIDTPTRAT = .CellValue(i, "IDTPTRATAMENTO") Then
                  ExibirAviso "Área já selecionada."
                  eResult = igEditResCancel
                  Exit For
               End If
            Next
         End If
         If .ColKey(lCol) = "IDTPTRATAMENTO" Then
            nIDAREA = xVal(.CellValue(lRow, "IDAREA"))
            For i = 1 To .RowCount - 2
               If nIDAREA = .CellValue(i, "IDAREA") And vNewValue = .CellValue(i, "IDTPTRATAMENTO") Then
                  ExibirAviso "Tratamento e Área já selecionados."
                  eResult = igEditResCancel
                  Exit For
               End If
            Next
         End If
      End If
      
      If eResult = igEditResCommit Then
         If .ColKey(lCol) = "IDAREA" Or .ColKey(lCol) = "IDTPTRATAMENTO" Or .ColKey(lCol) = "IDTPSERVICO" Then
            nTPSERV = xVal(.CellValue(lRow, "IDTPSERVICO"))
            nIDTPTRAT = xVal(.CellValue(lRow, "IDTPTRATAMENTO"))
            nIDAREA = xVal(.CellValue(lRow, "IDAREA"))
            
            If .ColKey(lCol) = "IDTPSERVICO" Then nTPSERV = xVal(vNewValue)
            If .ColKey(lCol) = "IDTPTRATAMENTO" Then nIDTPTRAT = xVal(vNewValue)
            If .ColKey(lCol) = "IDAREA" Then nIDAREA = xVal(vNewValue)
           
            Call SetNUMSESSAO(nTPSERV, nIDTPTRAT, nIDAREA, nNUMSESSAO, nIDDIRECAO, nFOTOTIPO, pDTATEND:=mvarMe.CmbDTATEND.Value)
            .CellValue(lRow, "NUMSESSAO") = nNUMSESSAO
            If bIsFOTO Then
               .CellValue(lRow, "IDDIRECAO") = nIDDIRECAO
               .CellValue(lRow, "FOTOTIPO") = nFOTOTIPO
            End If
            
'            Dim TbSessao As Object 'TB_OSESSAO
'            Set TbSessao = CriarObjeto("BANCO_3R.TB_OSESSAO")
'            Set TbSessao.xdb = mvarSys.xdb
'            If TbSessao.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO, Ch_IDSESSAO:=xVal(.CellValue(lRow, "IDSESSAO"))) Then
'               If TbSessao.IDAREA = xVal(vNewValue) Then
'                  .CellValue(lRow, "NUMSESSAO") = TbSessao.NUMSESSAO
'                  .CellValue(lRow, "IDDIRECAO") = TbSessao.IDDIRECAO
'               End If
'            Else
'            End If
         End If
      End If
   End With
End Sub
Private Sub SetNUMSESSAO(pIDTPSERVICO As Long, pIDTPTRATAMENTO As Long, pIDAREA As Long, ByRef pNUMSESSAO As Long, ByRef pIDDIRECAO As Long, ByRef pFOTOTIPO As Integer, Optional pDTATEND)
   Dim Sql As String
   
   pNUMSESSAO = 1
   pIDDIRECAO = 1
   If pIDTPSERVICO <> 1 Then '* Avaliação
      Sql = "Select IsNull(S.NUMSESSAO,0) [NUMSESSAO], IDDIRECAO, FOTOTIPO" & vbNewLine
      Sql = Sql & " From OSESSAO S " & vbNewLine
      Sql = Sql & " Join OATENDIMENTO A On S.IDLOJA =A.IDLOJA And S.IDATENDIMENTO =A.IDATENDIMENTO" & vbNewLine
      Sql = Sql & " Where A.IDLOJA=" & mvarIDLOJA & vbNewLine
      Sql = Sql & " And A.IDCLIENTE=" & mvarIDCLIENTE & vbNewLine
      Sql = Sql & " And S.IDTPSERVICO=" & pIDTPSERVICO & vbNewLine
      Sql = Sql & " And S.IDTPTRATAMENTO=" & pIDTPTRATAMENTO & vbNewLine
      If mvarSys.Propriedades("TPAREA") = "1" Or mvarSys.Propriedades("TPAREA") = "" Then
         Sql = Sql & " And ISNULL(S.IDAREA,0)=" & pIDAREA & vbNewLine
      End If
      If Not IsMissing(pDTATEND) Then
         Sql = Sql & " And A.DTATEND<" & SqlDate(pDTATEND) & vbNewLine
      End If
      Sql = Sql & " Order By 1"
      If mvarSys.xDb.AbreTabela(Sql) Then
         mvarSys.xDb.RsAux.MoveLast
         pNUMSESSAO = Val(mvarSys.xDb.RsAux("NUMSESSAO") & "") + 1
         pIDDIRECAO = IIf(mvarSys.xDb.RsAux("IDDIRECAO") = 2, 1, 2)
         pFOTOTIPO = xVal(mvarSys.xDb.RsAux("FOTOTIPO") & "")
      End If
   End If
End Sub
Private Sub mvarMe_GrdSESSAOColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   bDoDefault = False
End Sub
Private Sub mvarMe_GrdSESSAOColHeaderDblClick(ByVal lCol As Long)
   Dim i As Integer
   With mvarMe.GrdSESSAO
      If .ColKey(lCol) = "FOTOTIPO" Then
         For i = 1 To .RowCount - 1
            .CellValue(i, "FOTOTIPO") = Val(GetTag(mvarMe.TxtNOME, "FOTOTIPO", "1"))
         Next
      End If
   End With
End Sub
Private Sub mvarMe_GrdSESSAODblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   Dim nSessao As Integer
   Dim TbSessao As Object
   Dim Sql As String
   
   With mvarMe.GrdSESSAO
      If (lRow = .RowCount) Then
         bRequestEdit = False
      Else
         If .ColIndex("NUMSESSAO") = lCol Then
            If mvarIDATENDIMENTO = 0 Then
               Call ExibirInformacao("N° da sessão não pode ser alterada." & vbNewLine & "Atendimento precisa ser salvo.")
            Else
               nSessao = xVal(InputBox("Entre com o número da sessão do cliente", "Nº Sessão", .CellValue(lRow, "NUMSESSAO")))
               If nSessao <> 0 And nSessao <> xVal(.CellValue(lRow, "NUMSESSAO")) Then
                  
                  Sql = "Select S.*"
                  Sql = Sql & " From OSESSAO S "
                  Sql = Sql & " Join OATENDIMENTO A On A.IDLOJA=S.IDLOJA And A.IDATENDIMENTO=S.IDATENDIMENTO"
                  Sql = Sql & " Where A.IDLOJA=" & mvarIDLOJA
                  Sql = Sql & " And A.IDCLIENTE=" & mvarIDCLIENTE
                  Sql = Sql & " And S.IDTPSERVICO=" & .CellValue(lRow, "IDTPSERVICO")
                  Sql = Sql & " And S.IDTPTRATAMENTO=" & .CellValue(lRow, "IDTPTRATAMENTO")
                  Sql = Sql & " And S.IDAREA =" & .CellValue(lRow, "IDAREA")
                  Sql = Sql & " And S.NUMSESSAO =" & SqlNum(nSessao)
                  If mvarSys.xDb.ExisteReg(Sql) Then
                     Call ExibirInformacao("N° da sessão não pode ser alterada." & vbNewLine & "Já existe a " & nSessao & "º sessão para esta área.")
                  Else
                     Set TbSessao = CriarObjeto("BANCO_3R.TB_OSESSAO")
                     Set TbSessao.xDb = mvarSys.xDb
                     If TbSessao.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO, Ch_IDSESSAO:=xVal(.CellValue(lRow, "IDSESSAO"))) Then
                        TbSessao.NUMSESSAO = nSessao
                        If TbSessao.isDirt Then
                           If TbSessao.Salvar Then
                              .CellValue(lRow, "NUMSESSAO") = nSessao & "º"
                           End If
                        End If
                     End If
                     Set TbSessao = Nothing
                  End If
               End If
            End If
         End If
      End If
   End With
End Sub

Private Sub mvarMe_GrdSESSAOLostFocus()
   Dim sTag As String
   Dim i    As Integer
   Dim j    As Integer
   
   With mvarMe.GrdSESSAO
      '.Tag = SetTag(.Tag, "ISDIRT", 0)
      If .RowCount <= 1 Then .Tag = SetTag(.Tag, "ISDIRT", 1)
      For i = 1 To .RowCount - 1
         If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
         For j = 1 To .ColCount
            If GetTag(sTag, .ColKey(j), 0) <> .CellValue(i, j) Then
               .Tag = SetTag(.Tag, "ISDIRT", 1)
            End If
            If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
         Next
         If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
      Next
   End With
End Sub
Private Sub mvarMe_GrdSESSAOMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
   With mvarMe.GrdSESSAO
      If lRow = .RowCount Then
         If bVendaFechada Then
            Call ExibirInformacao("Venda fechada.")
            Exit Sub
         End If
         
         If lRow > 1 Then
            If Val(.CellValue(lRow - 1, "IDSESSAO")) = 0 Then Exit Sub
            If Val(.CellValue(lRow - 1, "IDTPSERVICO")) = 0 Then
               .CellValue(lRow - 1, "IDTPSERVICO") = 2
            End If
            If .ColVisible("IDAREA") Then
               If Val(.CellValue(lRow - 1, "IDTPTRATAMENTO")) = 0 Then Exit Sub
               If Val(.CellValue(lRow - 1, "IDAREA")) = 0 Then Exit Sub
            End If
            'If Val(.CellValue(lRow - 1, "IDDIRECAO")) = 0 Then Exit Sub
            'If Val(.CellValue(lRow - 1, "IDMANIPULO")) = 0 Then Exit Sub
         End If
'         If .ColVisible("IDAREA") Then
            Call IncluiGrdLinhaSessao
            If Not bIsFOTO Then .CellValue(lRow, "IDTPTRATAMENTO") = -1
'         Else
'            Call AdicionarTratamento
'         End If
         
      End If
   End With
End Sub
Private Sub mvarMe_GrdSESSAORequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   With mvarMe.GrdSESSAO
      bCancel = (GetTag(.ColTag(lCol), "EDIT", True) = False)
      If .CellFmtString(lRow, lCol) <> "" Then sText = .CellValue(lRow, lCol)
      eTextEditOpt = GetTag(.ColTag(lCol), "EDITOPT", 0)
   End With
End Sub
Private Sub mvarMe_GrdSESSAOValidate(Cancel As Boolean)
   mvarMe.GrdSESSAO.CommitEdit
End Sub
Private Sub mvarMe_GrdVendaColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   bDoDefault = False
End Sub
Private Sub mvarMe_GrdVendaDblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   If (lRow = mvarMe.GrdVenda.RowCount) Then
      bRequestEdit = False
   Else
      With mvarMe.GrdVenda
         If lRow > 0 Then
            If xVal(.CellValue(lRow, "IDVENDA")) = 0 Then
               Call ExibirVenda(-1)
               Call PopulaGRDVENDA
            Else
               Call mvarMe_CmdEditarClick
            End If
         End If
      End With
   End If
End Sub

Private Sub mvarMe_GrdVendaLostFocus()
   Dim sTag As String
   Dim i    As Integer
   Dim j    As Integer
   
   With mvarMe.GrdVenda
      '.Tag = SetTag(.Tag, "ISDIRT", 0)
      If .RowCount <= 1 Then .Tag = SetTag(.Tag, "ISDIRT", 1)
      For i = 1 To .RowCount - 1
         If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
         For j = 1 To .ColCount
            If GetTag(sTag, .ColKey(j), 0) <> .CellValue(i, j) Then .Tag = SetTag(.Tag, "ISDIRT", 1)
            If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
         Next
         If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
      Next
   End With
End Sub

Private Sub mvarMe_GrdVendaMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
   Dim bExiste As Boolean
   Dim nSaldo As Integer
   
   With mvarMe.GrdVenda
      If lRow = .RowCount Then
         If lRow > 1 Then
            If Val(.CellValue(lRow - 1, "IDVENDA")) = 0 Then Exit Sub
         End If
         bExiste = VerificarPacote(True, nSaldo)
         If nSaldo > 0 Or Not bExiste Then
            Call ExibirVenda(-1)
            Call PopulaGRDVENDA
         End If
      End If
   End With
End Sub
Private Sub mvarMe_GrdVendaRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   With mvarMe.GrdVenda
      bCancel = (GetTag(.ColTag(lCol), "EDIT", True) = False)
      If .CellFmtString(lRow, lCol) <> "" Then sText = .CellValue(lRow, lCol)
      eTextEditOpt = GetTag(.ColTag(lCol), "EDITOPT", 0)
   End With
End Sub
Private Sub mvarMe_GrdVendaValidate(Cancel As Boolean)
   mvarMe.GrdVenda.CommitEdit
End Sub

Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim lRow As Long
   Dim sMsg As String
   
   If KeyCode = vbKeyDelete Then
      If mvarMe.ActiveControl Is mvarMe.GrdSESSAO Or mvarMe.ActiveControl Is mvarMe.GrdPRODUTO Or mvarMe.ActiveControl Is mvarMe.GrdVenda Then
         With mvarMe.ActiveControl
            lRow = .CurRow
            If mvarMe.ActiveControl Is mvarMe.GrdVenda Then
               If bAtendFechado Then
                  Call ExibirAviso("Atendimento fechado não pode ser alterado.")
               Else
                  If .CurRow <= 0 Then Exit Sub
                  If xVal(.CellValue(.CurRow, "IDVENDA")) <> 0 Then
                     sMsg = "Esta operação irá desvincular a venda"
                     sMsg = sMsg & " '" & StrZero(.CellValue(.CurRow, "IDVENDA"), mvarSys.Propriedades("LENIDVENDA")) & "'"
                     sMsg = sMsg & " do atendimento"
                     sMsg = sMsg & " '" & StrZero(mvarIDATENDIMENTO, mvarSys.Propriedades("LENIDVENDA")) & "'."
                     sMsg = sMsg & vbNewLine & vbNewLine
                     sMsg = sMsg & "Deseja continuar?"
                     If vbNo = ExibirPergunta(sMsg, , False) Then
                        Exit Sub
                     End If
                  End If
         
                  If lRow <> .RowCount And lRow > 0 Then
                     .RemoveRow (.CurRow)
                     .Tag = SetTag(.Tag, "ISDIRT", 1)
                     .SetFocus
                     If lRow < .RowCount Then
                        .SetCurCell lRow, 1
                     ElseIf lRow > 1 Then
                        .SetCurCell lRow - 1, 1
                     End If
                  End If
                  Call SalvarVinculoVenda
               End If
            Else
               If lRow <> .RowCount And lRow > 0 Then
                  .RemoveRow (.CurRow)
                  .Tag = SetTag(.Tag, "ISDIRT", 1)
                  .SetFocus
                  If lRow < .RowCount Then
                     .SetCurCell lRow, 1
                  ElseIf lRow > 1 Then
                     .SetCurCell lRow - 1, 1
                  End If
               End If
            End If
            For lRow = 1 To .RowCount - 1
               .CellValue(lRow, 1) = lRow
            Next
         End With
      End If
   ElseIf KeyCode = vbKeyF4 Then
      Call mvarMe_CmdVendaClick
   ElseIf KeyCode = vbKeyF3 Then
      If mvarMe.ActiveControl Is mvarMe.TxtHHFIM Then
         mvarMe.CmbHHFIM.Value = mvarSys.xDb.Sysdate(2)
         mvarMe.TxtHHFIM.Text = Format(mvarMe.CmbHHFIM.Value, "hh:mm")
      End If
   ElseIf KeyCode >= vbKeyA And KeyCode <= vbKeyZ Then
      If mvarMe.ActiveControl Is mvarMe.TxtHHFIM Then
         If KeyCode = vbKeyN Then
            mvarMe.CmbHHFIM.Value = mvarSys.xDb.Sysdate(2)
            mvarMe.TxtHHFIM.Text = Format(mvarMe.CmbHHFIM.Value, "hh:mm")
         Else
            KeyCode = 0
         End If
      End If
   
   ElseIf KeyCode = vbKeySpace Then
      If mvarMe.ActiveControl Is mvarMe.TxtHHFIM Then
         mvarMe.CmbHHFIM.Value = Null
         mvarMe.TxtHHFIM.Enabled = False
         mvarMe.CmbHHFIM.SetFocus
      End If
   End If
End Sub

Private Sub mvarMe_Load()
   mvarMe.Icon = mvarSys.MDI.Icon
   
   bIsFOTO = IsFoto(mvarIDLOJA, mvarIDEVENTO)
   'If mvarSys.IDCOLIGADA = 2 Or mvarSys.IDCOLIGADA = 3 Then
   '   bIsFOTO = False
   'End If
   
   '* Monta Tela
   Call MontaTela
   
   '* Popula Tela
   Call CarregaFechamento
   Call PopulaTela
   
   DoEvents
   Screen.MousePointer = vbDefault
End Sub
Private Function IsFoto(pIDLOJA As Integer, pIDEVENTO As Long) As Boolean
   Dim sIDATEND   As Long
   Dim Sql As String
   Dim MyRS As Object
   
   If mvarSys.Propriedades("TPAREA") = "0" Then
      IsFoto = False
      Exit Function
   End If
   If pIDEVENTO <> 0 Then
      Sql = Sql_OATENDIMENTO(pIDLOJA, pIDEVENTO:=pIDEVENTO)
      If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
         sIDATEND = MyRS("IDATENDIMENTO")
      End If
      
      If sIDATEND <> 0 Then
         Sql = Sql_OSESSAO(pIDLOJA, pIDATENDIMENTO:=sIDATEND)
         Sql = Sql & "Order By IDSESSAO"
         If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
            nIDTRAT = MyRS("IDTPTRATAMENTO")
            
            Sql = Sql_OTPTRATAMENTO(pIDLOJA, nIDTRAT)
            If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
               If InArray(Trim(UCase(MyRS("DSCTRATAMENTO"))), Array("DEPILAÇÃO", "REJUVENESCIMENTO", "MANCHA", "ACNE", "FOTODEPILAÇÃO", "FOTOREJUVENESCIMENTO")) Then
                  IsFoto = True
               End If
            End If
         End If
      Else
         Sql = Sql_OSERVICOEVT(pIDLOJA, pIDEVENTO:=pIDEVENTO)
         Sql = Sql & "Order By IDSERVICOEVT"
         If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
            nIDTRAT = MyRS("IDTPTRATAMENTO")
            
            Sql = Sql_OTPTRATAMENTO(pIDLOJA, nIDTRAT)
            If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
               If InArray(Trim(UCase(MyRS("DSCTRATAMENTO"))), Array("DEPILAÇÃO", "REJUVENESCIMENTO", "MANCHA", "ACNE", "FOTODEPILAÇÃO", "FOTOREJUVENESCIMENTO")) Then
                  IsFoto = True
               End If
            End If
         End If
      End If
   End If
End Function
Private Sub MontaTela()
   mvarMe.TabAtend.Item(0).Selected = True
   mvarMe.TxtTEL1.MaxLength = mvarSys.Propriedades("MaxLenTel")
   Call SetTag(mvarMe.CmbHHFIM, "NULL", 1)
   
   Call MontaCmbFuncionario
   Call MontaCmbSala
   Call MontaCmbMaquina
   Call MontaGrdServico
   Call MontaGrdProduto
   Call MontaGrdVenda
End Sub
Private Sub PopulaObservacao()
   Dim Sql  As String
   Dim sObs As String
   Dim MyRS As Object
   
   sObs = ""
   
   Sql = "Select OBS "
   Sql = Sql & " From OCLIENTE"
   Sql = Sql & " Where IDLOJA = " & mvarIDLOJA
   Sql = Sql & " And IDCLIENTE = " & mvarIDCLIENTE
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      If Trim(MyRS("OBS") & "") <> "" Then
         sObs = sObs & "Cliente : " & MyRS("OBS") & "" & vbNewLine & vbNewLine
      End If
   End If
   
   
   Sql = "Select DTATEND, OBS "
   Sql = Sql & " From OATENDIMENTO "
   Sql = Sql & " Where IDLOJA = " & mvarIDLOJA
   Sql = Sql & " And IDCLIENTE = " & mvarIDCLIENTE
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      While Not MyRS.EOF
         If Trim(MyRS("OBS") & "") <> "" Then
            sObs = sObs & "Atendimento (" & Format(MyRS("DTATEND"), "dd/mm/yy") & "): " & MyRS("OBS") & "" & vbNewLine & vbNewLine
         End If
         MyRS.MoveNext
      Wend
   End If
   mvarMe.TxtObsOutras.Text = sObs
   
End Sub
Private Sub PopulaTela()
   Dim Sql     As String
   Dim TbEvt   As Object 'TB_OEVENTOAGENDA
   Dim TBFunc  As Object 'TB_RFUNCIONARIO
   Dim TBMaq   As Object 'TB_OMAQUINA
   Dim TbSala  As Object 'TB_OSALA
   
   Call LimparTela
  
   Set mvarMe.CmdFechar.Picture = mvarMe.lstImage.ListImages(IIf(bAtendFechado, 2, 1)).Picture
   Set mvarMe.CmdVenda.Picture = mvarMe.lstImage.ListImages(IIf(bVendaFechada, 4, 3)).Picture
   mvarMe.CmdOk.Enabled = Not bAtendFechado
   mvarMe.CmdSalvar.Enabled = Not bAtendFechado
   mvarMe.CmdChave.Visible = (bAtendFechado Or bVendaFechada)
   
   If mvarIDEVENTO <> 0 Then
      Sql = "IDEVENTO=" & mvarIDEVENTO
      If mvarTbAtend Is Nothing Then Set mvarTbAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
      Set mvarTbAtend.xDb = mvarSys.xDb
      If mvarTbAtend.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=Sql) Then
         With mvarTbAtend
            mvarIDATENDIMENTO = .IDATENDIMENTO
            mvarIDCLIENTE = .IDCLIENTE
            mvarMe.Caption = "Atendimento Nº" & StrZero(mvarIDATENDIMENTO, mvarSys.Propriedades("LENIDVENDA"))
            mvarMe.TxtIDATENDIMENTO.Text = StrZero(mvarIDATENDIMENTO, mvarSys.Propriedades("LENIDVENDA"))
            
            Call PopulaCliente(mvarIDCLIENTE)
            If LocalizarCombo(mvarMe.CmbIDFUNCIONARIO, .IDFUNCIONARIO, True, True) < 0 Then
               Set TBFunc = CriarObjeto("BANCO_3R.TB_RFUNCIONARIO", False)
               Set TBFunc.xDb = mvarSys.xDb
               If TBFunc.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDFUNCIONARIO:=.IDFUNCIONARIO) Then
                  mvarMe.CmbIDFUNCIONARIO.AddItem TBFunc.NOME
                  mvarMe.CmbIDFUNCIONARIO.ItemData(mvarMe.CmbIDFUNCIONARIO.NewIndex) = .IDFUNCIONARIO
                  Call LocalizarCombo(mvarMe.CmbIDFUNCIONARIO, .IDFUNCIONARIO, True, True)
               End If
            End If
            If LocalizarCombo(mvarMe.CmbIDSALA, .IDSALA, True, True) < 0 Then
               Set TbSala = CriarObjeto("BANCO_3R.TB_OSALA", False)
               Set TbSala.xDb = mvarSys.xDb
               If TbSala.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDMAQUINA:=.IDSALA) Then
                  mvarMe.CmbIDSALA.AddItem "Sala " & TbSala.CODSALA
                  mvarMe.CmbIDSALA.ItemData(mvarMe.CmbIDSALA.NewIndex) = .IDSALA
                  Call LocalizarCombo(mvarMe.CmbIDSALA, .IDSALA, True, True)
               End If
            End If
            
            If LocalizarCombo(mvarMe.CmbIDMAQUINA, .IDMAQUINA, True, True) < 0 Then
               Set TBMaq = CriarObjeto("BANCO_3R.TB_OMAQUINA", False)
               Set TBMaq.xDb = mvarSys.xDb
               If TBMaq.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDMAQUINA:=.IDMAQUINA) Then
                  mvarMe.CmbIDMAQUINA.AddItem TBMaq.CODMAQUINA
                  mvarMe.CmbIDMAQUINA.ItemData(mvarMe.CmbIDMAQUINA.NewIndex) = .IDMAQUINA
                  Call LocalizarCombo(mvarMe.CmbIDMAQUINA, .IDMAQUINA, True, True)
                  mvarMe.CmbIDMAQUINA.Enabled = (mvarMe.CmbIDMAQUINA.ListCount > 1)
               End If
            End If
            
            mvarMe.CmbDTATEND.Value = .dtatend
            mvarMe.CmbHHINI.Value = .HHINI
            mvarMe.CmbHHFIM.Value = .HHFIM
            mvarMe.TxtHHINI.Text = Format(.HHINI, "hh:mm")
            mvarMe.TxtHHFIM.Text = Format(.HHFIM, "hh:mm")
            
            Call SetTag(mvarMe.CmbHHFIM, "NULL", 0)
            mvarMe.TxtOBS.Text = .OBS
            
            Call PopulaGRDSESSAO
            Call PopulaGRDPRODUTO
            Call PopulaGRDVENDA
            Call PopulaObservacao
         End With
      Else
         Set TbEvt = CriarObjeto("BANCO_3R.TB_OEVENTOAGENDA")
         Set TbEvt.xDb = mvarSys.xDb
         If TbEvt.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEVENTO:=mvarIDEVENTO) Then
            With TbEvt
               mvarIDCLIENTE = .IDCLIENTE
               Call PopulaCliente(mvarIDCLIENTE)
               If mvarMe.CmbIDFUNCIONARIO.ListCount > 1 Then mvarMe.CmbIDFUNCIONARIO.ListIndex = -1
               If mvarMe.CmbIDSALA.ListCount > 1 Then mvarMe.CmbIDSALA.ListIndex = -1
               If mvarMe.CmbIDMAQUINA.ListCount > 1 Then mvarMe.CmbIDMAQUINA.ListIndex = -1
               
               If mvarIDSALA <> 0 Then
                  If LocalizarCombo(mvarMe.CmbIDSALA, CStr(mvarIDSALA), True, True) >= 0 Then
                     Call LocalizarCombo(mvarMe.CmbIDMAQUINA, CStr(mvarIDMAQUINA), True, True)
                  End If
               End If
               mvarMe.CmbDTATEND.Value = .StartDateTime
               mvarMe.CmbHHINI.Value = Format(.StartDateTime, "HH:mm")
               mvarMe.TxtHHINI.Text = Format(mvarMe.CmbHHINI.Value, "hh:mm")
               mvarMe.CmbHHFIM.Value = Format(.EndDateTime, "HH:mm")
               mvarMe.TxtHHFIM.Text = Format(mvarMe.CmbHHFIM.Value, "hh:mm")
               mvarMe.CmbHHFIM.Value = Null
               Call SetTag(mvarMe.CmbHHFIM, "NULL", 1)
               
               Call PopulaGRDServicoEVT
            End With
         End If
         Set TbEvt = Nothing
      End If
   End If
   Call ExibeDuracao
   Call ExibeUltSessao
   Call ExibeUltProd
End Sub
Private Sub CarregaFechamento()
   Dim Sql  As String
   Dim MyRS As Object
   
   Sql = "Select A.IDATENDIMENTO, A.DTATEND, A.SITATEND, V.IDVENDA, V.SITVENDA" & vbNewLine
   Sql = Sql & ", F.IDFATURA" & vbNewLine
   Sql = Sql & " From OATENDIMENTO A" & vbNewLine
   Sql = Sql & " Left Join OATENDIMENTO_VENDA AV On A.IDLOJA=AV.IDLOJA And A.IDATENDIMENTO=AV.IDATENDIMENTO" & vbNewLine
   Sql = Sql & " Left Join CVENDA V On AV.IDLOJA=V.IDLOJA And AV.IDVENDA=V.IDVENDA" & vbNewLine
   Sql = Sql & " Left Join FFATURA F On A.IDLOJA=F.IDLOJA And A.IDATENDIMENTO=F.IDATENDIMENTO" & vbNewLine
   Sql = Sql & " Where A.IDEVENTO=" & xVal(mvarIDEVENTO) & vbNewLine
   Sql = Sql & " And A.IDLOJA=" & xVal(mvarIDLOJA)
   Sql = Sql & " And A.IDATENDIMENTO<>0"
   
   mvarMe.CmdFechar.Enabled = False
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      bAtendFechado = (MyRS("SITATEND") & "" = gSITATEND_Fechado)
      bVendaFechada = (MyRS("SITVENDA") & "" = gSITVENDA_Fechada)
      bExisteVenda = (xVal(MyRS("IDVENDA") & "") <> 0)
      If Not mvarCalEvent Is Nothing Then
         mvarCalEvent.CustomProperties("VENDAFECHADA") = ""
      End If
      mvarMe.CmdFechar.Enabled = Not bAtendFechado And (MyRS("DTATEND") & "" < CDate(Format(Now, "dd/mm/yyyy")))
   End If
End Sub
Private Sub PopulaGRDPRODUTO()
   Dim MyTb    As Object 'TB_OATENDIMENTO_PRODUTO
   Dim MyTbP   As Object ' TB_SPRODUTO
   Dim sTag    As String
   Dim i       As Integer
   Dim Sql     As String
   Dim sUnid   As String

   Set MyTbP = CriarObjeto("BANCO_3R.TB_SPRODUTO", False)
   Set MyTbP.xDb = mvarSys.xDb
   
   Set MyTb = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_PRODUTO", False)
   Set MyTb.xDb = mvarSys.xDb
      
   
   If MyTb.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO, Ch_ORDERBY:="IDPROD") Then
      With mvarMe.GrdPRODUTO
         .Clear
         Call IncluiGrdLinhaProd
         While Not MyTb.Rs.EOF
            MyTb.Popula
            sUnid = "Un"
            If MyTbP.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPROD:=MyTb.IDPROD) Then sUnid = MyTbP.UNIDCONTROLE
            
            Call IncluiGrdLinhaProd
         
            i = .RowCount - 1
            If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
            Call SetTag(sTag, "ITEM", i)
            Call SetTag(sTag, "IDPROD", MyTb.IDPROD)
            Call SetTag(sTag, "QTDCOMPRA", MyTb.QTDCOMPRA)
            Call SetTag(sTag, "UNIDCONTROLE", sUnid)
                        
            .RowTag(i) = sTag
            
            .CellValue(i, "ITEM") = i
            .CellValue(i, "IDPROD") = MyTb.IDPROD
            .CellValue(i, "QTDCOMPRA") = MyTb.QTDCOMPRA
            .CellValue(i, "UNIDCONTROLE") = sUnid
                     
            MyTb.Rs.MoveNext
         Wend
         .RowVisible(1) = True
         '.CellSelected(1, "FOTOTIPO") = True
      End With
   End If
   Set MyTb = Nothing
End Sub
Private Sub PopulaGRDVENDA()
   Dim MyTb    As Object 'TB_OATENDIMENTO_VENDA
   Dim MyTbV   As Object 'TB_CVENDA
   Dim TbMoeda As Object 'TB_FMOEDA
   Dim sTag    As String
   Dim i       As Integer
   Dim Sql     As String
   Dim sUnid   As String
   
   Set MyTbV = CriarObjeto("BANCO_3R.TB_CVENDA", False)
   Set MyTbV.xDb = mvarSys.xDb
   
   Set MyTb = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_VENDA", False)
   Set MyTb.xDb = mvarSys.xDb
      
   Set TbMoeda = CriarObjeto("BANCO_3R.TB_FMOEDA", False)
   Set TbMoeda.xDb = mvarSys.xDb
   
   If MyTb.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO) Then
      With mvarMe.GrdVenda
         .Clear
         Call IncluiGrdLinhaVenda
         While Not MyTb.Rs.EOF
            MyTb.Popula
                        
            Call MyTbV.Pesquisar(Ch_IDLOJA:=MyTb.IDLOJA, Ch_IDVENDA:=MyTb.IDVENDA)
            bChave = bChave Or (MyTbV.SITVENDA = gSITATEND_Fechado)
            If TbMoeda.IDMOEDA <> MyTbV.IDMOEDA Then
               Call TbMoeda.Pesquisar(Ch_IDMOEDA:=MyTbV.IDMOEDA)
            End If
            
            Call IncluiGrdLinhaVenda
         
            i = .RowCount - 1
            If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
            Call SetTag(sTag, "ITEM", i)
            Call SetTag(sTag, "IDVENDA", StrZero(MyTbV.IDVENDA, mvarSys.Propriedades("LENIDVENDA")))
            Call SetTag(sTag, "DTVENDA", MyTbV.DTVENDA)
            Call SetTag(sTag, "VLITENS", TbMoeda.SIGLAMOEDA & " " & ValBr(MyTbV.VLITENS))
            Call SetTag(sTag, "VLPGTO", TbMoeda.SIGLAMOEDA & " " & ValBr(MyTbV.VLPGTO))
                        
            .RowTag(i) = sTag
            
            .CellValue(i, "ITEM") = i
            .CellValue(i, "IDVENDA") = StrZero(MyTbV.IDVENDA, mvarSys.Propriedades("LENIDVENDA"))
            .CellValue(i, "DTVENDA") = Format(MyTbV.DTVENDA, "dd/mm/yyyy")
            .CellValue(i, "VLITENS") = TbMoeda.SIGLAMOEDA & " " & ValBr(MyTbV.VLITENS)
            .CellValue(i, "VLPGTO") = TbMoeda.SIGLAMOEDA & " " & ValBr(MyTbV.VLPGTO)
                     
            MyTb.Rs.MoveNext
         Wend
         .RowVisible(1) = True
         '.CellSelected(1, "FOTOTIPO") = True
         If .RowCount > 1 Then
            mvarMe.CmdEditar.Enabled = True
            Call .SetCurCell(1, 1)
         Else
            mvarMe.CmdEditar.Enabled = False
         End If
      End With
   Else
      mvarMe.GrdVenda.Clear
      Call IncluiGrdLinhaVenda
   End If
   Set MyTb = Nothing
End Sub
Private Sub PopulaGRDSESSAO()
   Dim MyTb As Object 'TB_OSESSAO
   Dim sTag As String
   Dim i    As Integer
   Dim Sql  As String

   Set MyTb = CriarObjeto("BANCO_3R.TB_OSESSAO", False)
   Set MyTb.xDb = mvarSys.xDb
      
   
   If MyTb.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO, Ch_ORDERBY:="IDSESSAO") Then
      With mvarMe.GrdSESSAO
         .Clear
         While Not MyTb.Rs.EOF
            MyTb.Popula
                  
            Call IncluiGrdLinhaSessao
         
            i = .RowCount - 1
            If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
            Call SetTag(sTag, "IDSESSAO", MyTb.IDSESSAO)
            Call SetTag(sTag, "IDTPSERVICO", MyTb.IDTPSERVICO)
            Call SetTag(sTag, "IDTPTRATAMENTO", MyTb.IDTPTRATAMENTO)
            Call SetTag(sTag, "IDAREA", MyTb.IDAREA)
            Call SetTag(sTag, "NUMSESSAO", MyTb.NUMSESSAO)
            Call SetTag(sTag, "IDDIRECAO", MyTb.IDDIRECAO)
            Call SetTag(sTag, "IDMANIPULO", MyTb.IDMANIPULO)
            Call SetTag(sTag, "FOTOTIPO", MyTb.FOTOTIPO)
            Call SetTag(sTag, "POTENCIA", MyTb.POTENCIA)
            Call SetTag(sTag, "DISPAROS", MyTb.DISPAROS)
            Call SetTag(sTag, "TEMPOSESSAO", MyTb.TEMPOSESSAO)
            'Call SetTag(sTag, "IDREACAOADV", MyTb.IDREACAOADV)
            
            .RowTag(i) = sTag
            
            .CellValue(i, "IDSESSAO") = MyTb.IDSESSAO
            .CellValue(i, "IDTPSERVICO") = MyTb.IDTPSERVICO
            .CellValue(i, "IDTPTRATAMENTO") = MyTb.IDTPTRATAMENTO
            .CellValue(i, "IDAREA") = MyTb.IDAREA
            .CellValue(i, "NUMSESSAO") = MyTb.NUMSESSAO
            
            If bIsFOTO Then
               .CellValue(i, "IDDIRECAO") = MyTb.IDDIRECAO
               .CellValue(i, "IDMANIPULO") = MyTb.IDMANIPULO
               .CellValue(i, "FOTOTIPO") = MyTb.FOTOTIPO
               '.CellValue(i, "IDTPMANIPULO") = 1
               .CellValue(i, "POTENCIA") = MyTb.POTENCIA
               .CellValue(i, "DISPAROS") = MyTb.DISPAROS
               .CellValue(i, "TEMPOSESSAO") = MyTb.TEMPOSESSAO
            End If
            MyTb.Rs.MoveNext
         Wend
         .RowVisible(1) = True
         '.CellSelected(1, "FOTOTIPO") = True
      End With
   End If
   Set MyTb = Nothing
End Sub
Private Sub PopulaGRDServicoEVT()
   Dim MyTb As Object 'TB_OSERVICOEVT
   Dim sTag As String
   Dim i    As Integer
   Dim Sql  As String
   
   Dim nNUMSESSAO As Long
   Dim nIDDIRECAO As Long
   Dim nFOTOTIPO  As Integer

   Set MyTb = CriarObjeto("BANCO_3R.TB_OSERVICOEVT")
   Set MyTb.xDb = mvarSys.xDb
   
   If MyTb.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEVENTO:=mvarIDEVENTO) Then
      With mvarMe.GrdSESSAO
         Call SetTag(.Tag, "ISDIRT", 1)
         .Clear
         While Not MyTb.Rs.EOF
            MyTb.Popula
                  
            Call IncluiGrdLinhaSessao
         
            i = .RowCount - 1
            If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
            Call SetTag(sTag, "IDSESSAO", MyTb.IDSERVICOEVT)
            Call SetTag(sTag, "IDTPSERVICO", MyTb.IDTPSERVICO)
            Call SetTag(sTag, "IDTPTRATAMENTO", MyTb.IDTPTRATAMENTO)
            Call SetTag(sTag, "IDAREA", MyTb.IDAREA)
            
            .RowTag(i) = sTag
            
            .CellValue(i, "IDSESSAO") = MyTb.IDSERVICOEVT
            .CellValue(i, "IDTPSERVICO") = MyTb.IDTPSERVICO
            .CellValue(i, "IDTPTRATAMENTO") = MyTb.IDTPTRATAMENTO
            .CellValue(i, "IDAREA") = MyTb.IDAREA
                        
            Call SetNUMSESSAO(MyTb.IDTPSERVICO, MyTb.IDTPTRATAMENTO, MyTb.IDAREA, nNUMSESSAO, nIDDIRECAO, nFOTOTIPO, pDTATEND:=mvarMe.CmbDTATEND.Value)
            .CellValue(i, "NUMSESSAO") = nNUMSESSAO
            If bIsFOTO Then
               .CellValue(i, "IDDIRECAO") = nIDDIRECAO
               If GetTag(mvarMe.TxtNOME, "FOTOTIPO", 0) <> 0 And nFOTOTIPO = 0 Then
                  .CellValue(i, "FOTOTIPO") = xVal(GetTag(mvarMe.TxtNOME, "FOTOTIPO", 0))
               Else
                  .CellValue(i, "FOTOTIPO") = nFOTOTIPO
               End If
               .CellValue(i, "POTENCIA") = 0
               .CellValue(i, "IDMANIPULO") = 1
               '.CellValue(i, "IDTPMANIPULO") = 1
            End If
            MyTb.Rs.MoveNext
         Wend
         .RowVisible(1) = True
         If bIsFOTO Then
            .CellSelected(1, "FOTOTIPO") = True
         Else
            .CellSelected(1, "IDAREA") = True
         End If
         .Tag = SetTag(.Tag, "ISDIRT", 1)
      End With
   Else
      If mvarMe.GrdSESSAO.RowCount = 2 Then
         mvarMe.GrdSESSAO.RemoveRow 1
      End If
   End If
   Set MyTb = Nothing
End Sub
Private Sub LimparTela()
   With mvarMe
      mvarIDCLIENTE = 0
      .TxtNOME.Tag = ""
      .TxtTEL1.Text = ""
      .TxtNOME.Text = ""
      .TxtFOTOTIPO.Text = ""
      .TxtOBS.Text = ""
      If .CmbIDFUNCIONARIO.ListCount > 1 Then .CmbIDFUNCIONARIO.ListIndex = -1
      If .CmbIDSALA.ListCount > 1 Then .CmbIDSALA.ListIndex = -1
      If .CmbIDMAQUINA.ListCount > 1 Then .CmbIDMAQUINA.ListIndex = -1
      .CmbDTATEND.Value = mvarSys.xDb.Sysdate()
      .CmbHHINI.Value = mvarSys.xDb.Sysdate(2)
      .CmbHHFIM.Value = .CmbHHINI.Value
      .TxtHHINI.Text = Format(mvarMe.CmbHHINI, "hh:mm")
      .TxtHHFIM.Text = Format(mvarMe.CmbHHFIM, "hh:mm")
      
      .CmbHHFIM.Value = Null
      Call SetTag(.CmbHHFIM, "NULL", 1)
      .GrdSESSAO.Clear False
      Call IncluiGrdLinhaSessao
      .GrdPRODUTO.Clear False
      Call IncluiGrdLinhaProd
   End With
End Sub
Private Sub PopulaCliente(Optional pIDCLIENTE As Long, Optional pTelefone As String = "")
   Dim bAchou  As Boolean
   Dim Sql     As String
   Dim MyRS    As Object
   
   Sql = Sql_OCLIENTE(mvarIDLOJA)
   If IsMissing(pIDCLIENTE) Then
      Sql = Sql & " Where Tel1='" & pTelefone & "'"
      If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
         bAchou = True
      Else
         Sql = Sql & " Where Tel2='" & pTelefone & "'"
         If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
            bAchou = True
         End If
      End If
   Else
      Sql = Sql_OCLIENTE(mvarIDLOJA, pIDCLIENTE)
      If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
         bAchou = True
      End If
   End If
   If bAchou Then
      mvarIDCLIENTE = MyRS("IDCLIENTE")
      mvarMe.TxtIDCLIENTE.Text = MyRS("IDCLIENTE")
      mvarMe.TxtNOME.Text = MyRS("NOME")
      mvarMe.TxtTEL1.Text = IIf(MyRS("TEL1") & "" = "", MyRS("TEL2") & "", MyRS("TEL1") & "")
      mvarMe.TxtFOTOTIPO.Text = GetFOTOTIPOText(MyRS("FOTOTIPO") & "")
      If Trim(MyRS("OBS") & "") <> "" Then
         mvarMe.CmdIDCLIENTE.Caption = "Nome *"
      End If
      
      Call SetTag(mvarMe.TxtNOME, "IDCLIENTE", MyRS("IDCLIENTE") & "")
      Call SetTag(mvarMe.TxtNOME, "FOTOTIPO", MyRS("FOTOTIPO") & "")
      Call SetTag(mvarMe.TxtNOME, "TEL1", MyRS("TEL1") & "")
      Call SetTag(mvarMe.TxtNOME, "TEL2", MyRS("TEL2") & "")
      Call SetTag(mvarMe.TxtNOME, "EMAIL", MyRS("EMAIL") & "")
   Else
      mvarIDCLIENTE = 0
      mvarMe.TxtNOME.Tag = ""
      mvarMe.TxtIDCLIENTE.Text = ""
      mvarMe.TxtNOME.Text = ""
      mvarMe.TxtTEL1.Text = ""
      mvarMe.TxtFOTOTIPO.Text = ""
   End If
End Sub

Private Sub MontaCmbFuncionario()
   Dim TbCmb  As Object 'TB_RFUNCIONARIO
   Dim oCombo As XtremeSuiteControls.ComboBox
   Dim sWhere As String
      
   Set oCombo = mvarMe.CmbIDFUNCIONARIO
   
   Set TbCmb = CriarObjeto("BANCO_3R.TB_RFUNCIONARIO", False)
   With TbCmb
      Set .xDb = mvarSys.xDb
      oCombo.Clear
      sWhere = ""
      sWhere = sWhere & " SITFUNC='A'"
      sWhere = sWhere & " And FLGCERTIFICADO=1"
      
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=sWhere) Then
         While Not .Rs.EOF
            .Popula
            oCombo.AddItem .NOME
            oCombo.ItemData(oCombo.NewIndex) = .IDFUNCIONARIO
            .Rs.MoveNext
         Wend
      Else
         Call ExibirInformacao("Favor cadastrar ao menos um Funcionário.", "Cadastro")
         Dim TlCad As Object
         Set TlCad = CriarObjeto("CADASTRO3R.TL_CADRFUNCIONARIO", False)
         Set TlCad.Sys = mvarSys
         TlCad.Show
         If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=sWhere) Then
            Call MontaCmbFuncionario
         Else
            Exit Sub
         End If
         Exit Sub
      End If
   End With
   Set TbCmb = Nothing
   With oCombo
      .ListIndex = IIf(.ListCount = 1, 0, -1)
      .Enabled = (.ListCount > 1)
   End With
End Sub
Private Sub MontaCmbSala()
   Dim TbCmb  As Object
   Dim oCombo As XtremeSuiteControls.ComboBox
      
   Set oCombo = mvarMe.CmbIDSALA
   Set TbCmb = CriarObjeto("BANCO_3R.TB_OSALA", False)
   With TbCmb
      Set .xDb = mvarSys.xDb
      oCombo.Clear
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="ATIVO=1") Then
         While Not .Rs.EOF
            .Popula
            oCombo.AddItem "Sala " & .CODSALA
            oCombo.ItemData(oCombo.NewIndex) = .IDSALA
            .Rs.MoveNext
         Wend
      End If
   End With
   Set TbCmb = Nothing
   With oCombo
      .ListIndex = IIf(.ListCount = 1, 0, -1)
      If .ListIndex < 0 Then
         Call LocalizarCombo(oCombo, CStr(mvarIDSALA), True, True)
      End If
      .Enabled = (.ListCount > 1)
   End With
End Sub
Private Sub MontaCmbMaquina()
   Dim TbCmb  As Object 'TB_OMAQUINA
   Dim oCombo As XtremeSuiteControls.ComboBox
   Dim Sql    As String
      
   Set oCombo = mvarMe.CmbIDMAQUINA
   
   Set TbCmb = CriarObjeto("BANCO_3R.TB_OMAQUINA", False)
   With TbCmb
      'If mvarIDSALA <= 0 Then mvarIDSALA = mvarMe.CmbIDSALA.ItemData(mvarMe.CmbIDSALA.ListIndex)
      'If mvarIDSALA <= 0 Then mvarIDSALA = 1

      Set .xDb = mvarSys.xDb
      oCombo.Clear
      Sql = "Select Distinct M.* "
      Sql = Sql & " From OMAQUINA M Join OSALA_MAQUINA S On M.IDLOJA=S.IDLOJA And M.IDMAQUINA=S.IDMAQUINA"
      Sql = Sql & " Where M.SITMAQUINA=1"
      Sql = Sql & " And S.IDLOJA=" & mvarIDLOJA
'      Sql = Sql & " And S.IDSALA=" & mvarIDSALA  'mvarMe.CmbIDSALA.ItemData(mvarMe.CmbIDSALA.ListIndex)
      Sql = Sql & " And S.DTINICIO <=" & SqlDate(Format(Now(), "dd/mm/yyyy") & " 23:59")
      Sql = Sql & " And (S.DTFIM is Null Or S.DTFIM >=" & SqlDate(Format(Now(), "dd/mm/yyyy") & " 00:00") & ")"
      If mvarMe.CmbIDSALA.ListIndex >= 0 Then
         Sql = Sql & " And S.IDSALA=" & mvarMe.CmbIDSALA.ItemData(mvarMe.CmbIDSALA.ListIndex)
      End If
      Sql = Sql & " Order By M.CODMAQUINA"
            
      If mvarSys.xDb.AbreTabela(Sql) Then
         If mvarIDATENDIMENTO = 0 Then oCombo.Clear
         While Not mvarSys.xDb.RsAux.EOF
            .Popula mvarSys.xDb.RsAux
            oCombo.AddItem .CODMAQUINA
            oCombo.ItemData(oCombo.NewIndex) = .IDMAQUINA
            mvarSys.xDb.RsAux.MoveNext
         Wend
      Else
         Sql = "Select S.IDEVENTO" & vbNewLine
         Sql = Sql & " From OSERVICOEVT S " & vbNewLine
         Sql = Sql & " Join OTPTRATAMENTO T On S.IDLOJA=T.IDLOJA And S.IDTPTRATAMENTO=T.IDTPTRATAMENTO" & vbNewLine
         Sql = Sql & " Where S.IDLOJA=" & mvarIDLOJA & vbNewLine
         Sql = Sql & " And S.IDEVENTO=" & mvarIDEVENTO & vbNewLine
         Sql = Sql & " And T.FLGDISPARO=1"
         bFLGDISPARO = mvarSys.xDb.ExisteReg(Sql)
         If bFLGDISPARO Then
            Call ExibirInformacao("Favor cadastrar máquina deste tratamento para esta sala.", "Cadastro")
            Dim TlCad As Object
            Set TlCad = CriarObjeto("CADASTRO3R.TL_CADOMAQ", False)
            Set TlCad.Sys = mvarSys
            TlCad.Show
            If mvarSys.xDb.AbreTabela(Sql) Then
               Sql = "Select S.IDEVENTO" & vbNewLine
               Sql = Sql & " From OSERVICOEVT S " & vbNewLine
               Sql = Sql & " Join OTPTRATAMENTO T On S.IDLOJA=T.IDLOJA And S.IDTPTRATAMENTO=T.IDTPTRATAMENTO" & vbNewLine
               Sql = Sql & " Where S.IDLOJA=" & mvarIDLOJA & vbNewLine
               Sql = Sql & " And S.IDEVENTO=" & mvarIDEVENTO & vbNewLine
               Sql = Sql & " And T.FLGDISPARO=1"
               bFLGDISPARO = mvarSys.xDb.ExisteReg(Sql)
               If bFLGDISPARO Then
                  Call ExibirInformacao("Este tratamento necessita de uma máquina nesta sala." & vbNewLine & "Favor associar máquina à sala.")
               Else
                  Call MontaCmbMaquina
               End If
            End If
            Exit Sub
         Else
            mvarMe.CmbIDMAQUINA.Enabled = False
         End If
      End If
   End With
   Set TbCmb = Nothing
   With oCombo
      .ListIndex = IIf(.ListCount = 1, 0, -1)
      .Enabled = (.ListCount > 1)
   End With
End Sub
Private Sub MontaGrdServico()
   Dim MyRS       As Object
   Dim Sql        As String
   Dim TbDirecao  As Object 'TB_ODIRECAO
   Dim TbFoto     As Object 'TB_OFOTOTIPO
   Dim sAux       As String
   Dim i          As Integer
   
   On Error Resume Next
   If bIsFOTO Then
      Set TbDirecao = CriarObjeto("BANCO_3R.TB_ODIRECAO", False)
      Set TbDirecao.xDb = mvarSys.xDb
      Set TbFoto = CriarObjeto("BANCO_3R.TB_OFOTOTIPO", False)
      Set TbFoto.xDb = mvarSys.xDb
   End If
   With mvarMe.GrdSESSAO
      .ShowControlsInAllCells = True
      .Header.DragCols = False
      .SilentValidation = True
      '.Header.ImageList = mvarMe.IlsIcons
      '.ImageList = mvarMe.IlsIcons
      With .Combos
         With .Add("IDTPSERVICO")
            Sql = ""
            Sql = Sql & "Select Distinct T.IDTPSERVICO, T.DSCSERVICO"
            Sql = Sql & " From OTPSERVICO T"
            Sql = Sql & " Left Join OSESSAO S On S.IDLOJA=T.IDLOJA And S.IDTPSERVICO=T.IDTPSERVICO and S.IDATENDIMENTO=" & mvarIDATENDIMENTO
            Sql = Sql & " Where T.IDLOJA=" & mvarIDLOJA
            Sql = Sql & " And (T.ATIVO=1 OR NOT S.IDTPSERVICO IS NULL)"
            Sql = Sql & " Order by T.IDTPSERVICO"
            If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
               While Not MyRS.EOF
                  .AddItem MyRS("DSCSERVICO"), xVal(MyRS("IDTPSERVICO"))
                  MyRS.MoveNext
               Wend
            End If
            Set MyRS = Nothing
            .AutoAdjustWidth
         End With
         With .Add("IDTPTRATAMENTO")
            Sql = ""
            Sql = Sql & "Select Distinct T.IDTPTRATAMENTO, T.DSCTRATAMENTO" & vbNewLine
            Sql = Sql & " From OTPTRATAMENTO T" & vbNewLine
            Sql = Sql & " Left Join OSESSAO S On S.IDLOJA=T.IDLOJA And S.IDTPTRATAMENTO=T.IDTPTRATAMENTO and S.IDATENDIMENTO=" & mvarIDATENDIMENTO & vbNewLine
            Sql = Sql & " Where T.IDLOJA=" & mvarIDLOJA & vbNewLine
            Sql = Sql & " And (T.ATIVO=1 OR NOT S.IDTPTRATAMENTO IS NULL)" & vbNewLine
            Sql = Sql & " Order by T.DSCTRATAMENTO"
            If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
               While Not MyRS.EOF
                  .AddItem MyRS("DSCTRATAMENTO"), xVal(MyRS("IDTPTRATAMENTO"))
                  MyRS.MoveNext
               Wend
            End If
            Set MyRS = Nothing
            .MaxHeightInItems = 10
            .AutoAdjustWidth
         End With
         With .Add("IDAREA")
            Sql = ""
            Sql = Sql & "Select Distinct A.IDAREA, A.DSCAREA"
            Sql = Sql & " From OAREA A"
            Sql = Sql & " Left Join OSESSAO S On S.IDLOJA=A.IDLOJA And S.IDAREA=A.IDAREA and S.IDATENDIMENTO=" & mvarIDATENDIMENTO
            Sql = Sql & " Where A.IDLOJA=" & mvarIDLOJA
            Sql = Sql & " And (A.ATIVO=1 OR NOT S.IDAREA IS NULL)"
            Sql = Sql & " Order by A.DSCAREA"
            If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
               While Not MyRS.EOF
                  .AddItem MyRS("DSCAREA"), xVal(MyRS("IDAREA"))
                  MyRS.MoveNext
               Wend
            End If
            Set MyRS = Nothing
            .MaxHeightInItems = 10
            .AutoAdjustWidth
         End With
         
         'bIsFOTO = False
         If bIsFOTO Then
            With .Add("FOTOTIPO")
               If TbFoto.Pesquisar(Ch_Where:="", Ch_ORDERBY:="FOTOTIPO") Then
                  While Not TbFoto.Rs.EOF
                     TbFoto.Popula
                     sAux = GetFOTOTIPOText(TbFoto.FOTOTIPO, TbFoto.PELE)
                     .AddItem sAux, TbFoto.FOTOTIPO
                     TbFoto.Rs.MoveNext
                  Wend
               End If
               Set TbFoto = Nothing
               .AutoAdjustWidth
            End With
            With .Add("IDDIRECAO")
               If TbDirecao.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_ORDERBY:="NORDEM") Then
                  While Not TbDirecao.Rs.EOF
                     TbDirecao.Popula
                     .AddItem TbDirecao.DSCDIRECAO, TbDirecao.IDDIRECAO
                     TbDirecao.Rs.MoveNext
                  Wend
               End If
               Set TbDirecao = Nothing
               .AutoAdjustWidth
            End With
         End If
      End With
      
      
      With .AddCol(sKey:="IDSESSAO", sHeader:="#", lWidth:=25, bVisible:=True)
         .eType = igCellText
         .sCtrlKey = "IDSESSAO"
         .bSelected = False
          
      End With
      With .AddCol(sKey:="IDTPSERVICO", sHeader:="Serviço", lWidth:=90, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDTPSERVICO"
      End With
      With .AddCol(sKey:="IDTPTRATAMENTO", sHeader:="Tratamento", lWidth:=IIf(bIsFOTO, 115, 150), bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDTPTRATAMENTO"
      End With
      With .AddCol(sKey:="IDAREA", sHeader:="Área", lWidth:=115, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDAREA"
      End With
      With .AddCol(sKey:="NUMSESSAO", sHeader:="Nº", lWidth:=25, bVisible:=True)
         .eType = igCellText
         .sCtrlKey = "NUMSESSAO"
         .eTextFlags = igTextCenter
         .sFmtString = "#º"
      End With
            
      With .AddCol(sKey:="IDDIRECAO", sHeader:="Direção", lWidth:=85, bVisible:=(mvarSys.Propriedades("TPDIR") = "1") And bIsFOTO)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDDIRECAO"
      End With
      With .AddCol(sKey:="FOTOTIPO", sHeader:="Fototipo", lWidth:=85, bVisible:=bIsFOTO)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "FOTOTIPO"
      End With
      With .AddCol(sKey:="IDMANIPULO", sHeader:="Manípulo", lWidth:=85, bVisible:=False)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDMANIPULO"
      End With
      'With .AddCol(sKey:="IDTPMANIPULO", sHeader:="TPMANIPULO", lWidth:=85, bVisible:=False)
      '   .eType = igCellCombo
      '   .eTypeFlags = igComboBtnFlat
      '   .sCtrlKey = "IDTPMANIPULO"
      'End With
      With .AddCol(sKey:="POTENCIA", sHeader:="Potência", lWidth:=60, bVisible:=bIsFOTO)
         .eType = igCellText
         .sCtrlKey = "POTENCIA"
         .eTextFlags = igTextCenter
      End With
      With .AddCol(sKey:="DISPAROS", sHeader:="Disparos", lWidth:=60, bVisible:=(mvarSys.Propriedades("TPDISP") = "1") And bIsFOTO)
         .eType = igCellText
         .sCtrlKey = "DISPAROS"
         .eTextFlags = igTextCenter
      End With
      With .AddCol(sKey:="TEMPOSESSAO", sHeader:="Tempo", lWidth:=50, bVisible:=False)
         .eType = igCellText
         .sCtrlKey = "TEMPOSESSAO"
         .eTextFlags = igTextCenter
      End With
   
      For i = 1 To .ColCount
         .ColTag(i) = ""
      Next
                       
      .ColTag("IDSESSAO") = SetTag(.ColTag("IDSESSAO"), "EDIT", False)
      .ColTag("NUMSESSAO") = SetTag(.ColTag("NUMSESSAO"), "EDIT", False)
      .ColTag("NUMSESSAO") = SetTag(.ColTag("NUMSESSAO"), "EDITOPT", igTextEditNumberOnly)
      
      If bIsFOTO Then
         .ColTag("POTENCIA") = SetTag(.ColTag("POTENCIA"), "EDITOPT", igTextEditNumberOnly)
         .ColTag("DISPAROS") = SetTag(.ColTag("DISPAROS"), "EDITOPT", igTextEditNumberOnly)
         .ColTag("TEMPOSESSAO") = SetTag(.ColTag("TEMPOSESSAO"), "EDITOPT", igTextEditNumberOnly)
      Else
         Dim Rs As Object
         Dim RsLista As Object
         Dim sItem As String
         Dim nWidth  As Long
         Dim sKey    As String
         
         Sql = "Select * "
         Sql = Sql & " From OTPTRATATR "
         Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
         Sql = Sql & " And IDTPTRATAMENTO=" & SqlNum(nIDTRAT)
         Sql = Sql & " And ATIVO=1"
         Sql = Sql & " Order by IDATR"
         If mvarSys.xDb.AbreTabela(Sql, Rs) Then
            While Not Rs.EOF
               sKey = UCase(Trim(Rs("DSCATR") & ""))
               nWidth = IIf(Len(sKey) * 8 < 50, 50, Len(sKey) * 8)
               If Trim(Rs("VLLISTA") & "") <> "" Then
                  If mvarSys.xDb.AbreTabela(Rs("VLLISTA"), RsLista) Then
                     .Combos.Add sKey
                     With .Combos(sKey)
                        While Not RsLista.EOF
                           If RsLista.Fields.Count <= 2 Then
                              sItem = CStr(RsLista.Fields(1).Value)
                           Else
                              sItem = CStr(RsLista.Fields(1).Value) & "-" & CStr(RsLista.Fields(2).Value)
                           End If
                           .AddItem sItem, RsLista.Fields(0).Value
                           RsLista.MoveNext
                        Wend
                        .AutoAdjustWidth
                        nWidth = .Width
                     End With
                  End If
               End If
            
               With .AddCol(sKey:=sKey, sHeader:=Rs("DSCATR"), lWidth:=nWidth, bVisible:=True)
                  If Trim(Rs("VLLISTA") & "") = "" Then
                     .eType = igCellText
                     .eTextFlags = igTextCenter
                  Else
                     .eType = igCellCombo
                     .eTypeFlags = igComboBtnFlat
                  End If
                  .sCtrlKey = sKey
               End With
               .ColTag(sKey) = ""
               .ColTag(sKey) = SetTag(.ColTag(sKey), "EDITOPT", IIf(Rs("TPATR") = "N", igTextEditNumberOnly, igTextEditUpperCase))
               Rs.MoveNext
            Wend
         End If
         Set Rs = Nothing
         
         'For i = 1 To .ColCount
         '   .ColTag(i) = ""
         'Next
         .ColTag("IDSESSAO") = ""
         .ColTag("NUMSESSAO") = ""
         .ColTag("NUMSESSAO") = ""
                          
         .ColTag("IDSESSAO") = SetTag(.ColTag("IDSESSAO"), "EDIT", False)
         .ColTag("NUMSESSAO") = SetTag(.ColTag("NUMSESSAO"), "EDIT", False)
         .ColTag("NUMSESSAO") = SetTag(.ColTag("NUMSESSAO"), "EDITOPT", igTextEditNumberOnly)
      
         .ColTag("IDTPSERVICO") = SetTag(CStr(.ColTag("IDTPSERVICO")), "EDIT", (.Combos("IDTPSERVICO").ListCount > 2) And (mvarSys.Propriedades("TPSERV", True) <> "0"))
         .ColTag("IDTPTRATAMENTO") = SetTag(CStr(.ColTag("IDTPTRATAMENTO")), "EDIT", True)
         .ColTag("IDAREA") = SetTag(CStr(.ColTag("IDAREA")), "EDIT", (.Combos("IDAREA").ListCount > 2) And (mvarSys.Propriedades("TPAREA", True) <> "0"))
      End If
      
      '************
      '* Configurar Estrutura do Serviço
      If mvarSys.Propriedades("TPSERV", True) = "0" Then
         If mvarSys.Propriedades("TPAREA", True) = "0" Then
            .ColHeaderText("IDTPSERVICO") = "Serviços"
            .ColHeaderText("IDTPTRATAMENTO") = ""
         End If
         .ColVisible("IDTPSERVICO") = True
         '.ColWidth("IDTPTRATAMENTO") = .ColWidth("IDTPTRATAMENTO") + .ColWidth("IDTPSERVICO")
      End If
      If mvarSys.Propriedades("TPAREA", True) = "0" Then
         .ColVisible("IDAREA") = False
         .ColWidth("IDTPTRATAMENTO") = .ColWidth("IDTPTRATAMENTO") + .ColWidth("IDAREA")
      End If
      If Not .ColVisible("IDDIRECAO") Then
         .ColWidth("IDTPTRATAMENTO") = .ColWidth("IDTPTRATAMENTO") + .ColWidth("IDDIRECAO")
      End If
      If Not .ColVisible("DISPAROS") Then
         .ColWidth("IDTPTRATAMENTO") = .ColWidth("IDTPTRATAMENTO") + .ColWidth("DISPAROS")
      End If
      'If Not .ColVisible("TEMPOSESSAO") Then
      '   .ColWidth("IDTPTRATAMENTO") = .ColWidth("IDTPTRATAMENTO") + .ColWidth("TEMPOSESSAO")
      'End If
      .ColVisible("IDTPTRATAMENTO") = (mvarSys.Propriedades("TPTRAT") = 1 Or mvarSys.Propriedades("TPTRAT") = "")
      
      .Combos("IDTPTRATAMENTO").Width = .ColWidth("IDTPTRATAMENTO")
      .Combos("IDTPTRATAMENTO").MaxHeightInItems = 10
                                   
      IncluiGrdLinhaSessao
   End With
End Sub
Private Sub IncluiGrdLinhaSessao()
   Dim lRow As Long
   Dim i As Integer
   Dim nColDefault As Integer
   
'   On Error Resume Next
   nColDefault = 4
   With mvarMe.GrdSESSAO
      .Header.Font.Size = 9
      .Header.Font.Name = mvarMe.Font.Name
      .Font.Size = .Header.Font.Size
      .Font.Name = .Header.Font.Name
      
      lRow = .RowCount
      If lRow > 0 Then .RemoveRow lRow
      If lRow > 1 Then
         If GetTag(.ColTag("IDTPSERVICO"), "EDIT", True) = False Then
            .CellForeColor(lRow - 1, "IDTPSERVICO") = vbGrayText
         Else
            .CellForeColor(lRow - 1, "IDTPSERVICO") = vbBlack
         End If
      End If

      .AddRow
      lRow = .RowCount
      .RowHeight(.RowCount) = 20

      .CellType(lRow, "IDSESSAO") = igCellText
      .CellValue(lRow, "IDSESSAO") = lRow
      
      .CellType(lRow, "IDTPSERVICO") = igCellCombo
      .CellValue(lRow, "IDTPSERVICO") = 2
      If lRow > 1 Then .CellValue(lRow, "IDTPSERVICO") = .CellValue(lRow - 1, "IDTPSERVICO")
      

      .CellType(lRow, "IDTPTRATAMENTO") = igCellCombo
      .CellValue(lRow, "IDTPTRATAMENTO") = 1
      If lRow > 1 Then .CellValue(lRow, "IDTPTRATAMENTO") = .CellValue(lRow - 1, "IDTPTRATAMENTO")

      .CellType(lRow, "IDAREA") = igCellCombo
      .CellValue(lRow, "IDAREA") = ""

      .CellType(lRow, "NUMSESSAO") = igCellText
      .CellValue(lRow, "NUMSESSAO") = ""

      If bIsFOTO Then
         .CellType(lRow, "IDDIRECAO") = igCellCombo
         .CellValue(lRow, "IDDIRECAO") = 1
         If lRow > 1 Then .CellValue(lRow, "IDDIRECAO") = .CellValue(lRow - 1, "IDDIRECAO")
   
         .CellType(lRow, "IDMANIPULO") = igCellText
         .CellValue(lRow, "IDMANIPULO") = 1
         .CellValue(lRow, "IDMANIPULO") = ""
         
         .CellType(lRow, "FOTOTIPO") = igCellCombo
         .CellValue(lRow, "FOTOTIPO") = ""
         If lRow >= 2 Then
            .CellValue(lRow, "FOTOTIPO") = .CellValue(lRow - 1, "FOTOTIPO")
         Else
            If GetTag(mvarMe.TxtNOME, "FOTOTIPO", 0) <> 0 Then
               .CellValue(lRow, "FOTOTIPO") = xVal(GetTag(mvarMe.TxtNOME, "FOTOTIPO", 0))
            End If
         End If
         
         .CellType(lRow, "POTENCIA") = igCellText
         .CellValue(lRow, "POTENCIA") = "0"
         
         .CellType(lRow, "DISPAROS") = igCellText
         .CellValue(lRow, "DISPAROS") = "0"
         
         .CellType(lRow, "TEMPOSESSAO") = igCellText
         .CellValue(lRow, "TEMPOSESSAO") = "0"
               
   '      .CellType(lRow, "IDREACAOADV") = igCellText
   '      .CellValue(lRow, "IDREACAOADV") = ""
      End If
      .CellForeColor(lRow, "IDTPSERVICO") = vbBlack
      
      For i = 1 To .ColCount
         If GetTag(.ColTag(i), "EDIT", True) = False Then
            .CellForeColor(lRow, i) = vbGrayText
         End If
      Next
      
            
      .AddRow
      
      For i = 1 To .ColCount
         .CellType(.RowCount, i) = igCellText
      Next
      For i = 1 To .ColCount
         If .ColVisible(i) Then
            .CellValue(.RowCount, i) = "Clique para incluir nova linha."
            .CellForeColor(.RowCount, i) = vbGrayText
            .CellTextFlags(.RowCount, i) = igTextNoClip Or igTextCenter
            Exit For
         End If
      Next
      .RowMode = (.RowCount = 1)
      .SetCurCell .RowCount, 1
      .SetCurCell lRow, nColDefault
      If .Visible Then
         .SetFocus
      End If
      .Redraw = True
   End With
End Sub
Private Sub MontaGrdProduto()
   Dim Sql        As String
   Dim sAux       As String
   Dim MyRS       As Object
   Dim i          As Integer
   
   On Error Resume Next
   With mvarMe.GrdPRODUTO
      .ShowControlsInAllCells = True
      .Header.DragCols = False
      .SilentValidation = True
      '.Header.ImageList = mvarMe.IlsIcons
      '.ImageList = mvarMe.IlsIcons
            
      With .Combos
         With .Add("IDPROD")
            Sql = ""
            Sql = Sql & "Select Distinct P.IDPROD, P.NMPROD"
            Sql = Sql & " From OPRODUTO P"
            Sql = Sql & " Left Join OSESSAO S On S.IDLOJA=P.IDLOJA And S.IDPROD=P.IDPROD And S.IDATENDIMENTO=" & mvarIDATENDIMENTO
            Sql = Sql & " Where P.IDLOJA=" & mvarIDLOJA
            Sql = Sql & " And P.ESERVICO=0"
            Sql = Sql & " And (P.SITPROD='A' OR NOT S.IDPROD IS NULL)"
            If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
               While Not MyRS.EOF
                  sAux = MyRS("NMPROD")
                  sAux = sAux & Space(65 - Len(sAux))
                  .AddItem sAux, MyRS("IDPROD")
                  MyRS.MoveNext
               Wend
            End If
            Set MyRS = Nothing
            .AutoAdjustWidth
         End With
      End With
      
      With .AddCol(sKey:="ITEM", sHeader:="#", lWidth:=25, bVisible:=True)
         .eType = igCellText
         .sCtrlKey = "ITEM"
      End With
      
      
      With .AddCol(sKey:="IDPROD", sHeader:="Produto", lWidth:=250, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDPROD"
      End With
      
      With .AddCol(sKey:="QTDCOMPRA", sHeader:="Quantidade", lWidth:=70, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextCenter
         .sCtrlKey = "QTDCOMPRA"
      End With
            
      With .AddCol(sKey:="UNIDCONTROLE", sHeader:="Unid.", lWidth:=40, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextCenter
         .sCtrlKey = "UNIDCONTROLE"
      End With
      
      For i = 1 To .ColCount
         .ColTag(i) = ""
      Next
      
      .ColTag("ITEM") = SetTag(.ColTag("ITEM"), "EDIT", False)
      .ColTag("IDPROD") = SetTag(.ColTag("IDPROD"), "EDIT", True)
      .ColTag("QTDCOMPRA") = SetTag(.ColTag("QTDCOMPRA"), "EDIT", True)
      .ColTag("UNIDCONTROLE") = SetTag(.ColTag("UNIDCONTROLE"), "EDIT", False)
            
      .ColTag("QTDCOMPRA") = SetTag(.ColTag("QTDCOMPRA"), "EDITOPT", igTextEditNumberOnly)

      IncluiGrdLinhaProd

   End With
End Sub
Private Sub IncluiGrdLinhaProd()
   Dim lRow As Long
   Dim i As Integer
   Dim nColDefault As Integer
         
   nColDefault = 2
   With mvarMe.GrdPRODUTO
      If .RowCount - 1 >= .Combos("IDPROD").ListCount Then
         If .Combos("IDPROD").ListCount = 0 Then
            Call ExibirAviso("Não existe produto de venda cadastrado.", "Produtos")
         End If
         Exit Sub
      End If
   
      .Header.Font.Size = 9
      .Header.Font.Name = mvarMe.Font.Name
      .Font.Size = .Header.Font.Size
      .Font.Name = .Header.Font.Name
      
      lRow = .RowCount
      If lRow > 0 Then .RemoveRow lRow
      If lRow > 1 Then .CellForeColor(lRow - 1, "IDPROD") = vbBlack
      
      If lRow <> 0 Then
         .AddRow
         lRow = .RowCount
         .RowHeight(.RowCount) = 20
   
         .CellType(lRow, "ITEM") = igCellText
         .CellValue(lRow, "ITEM") = lRow
         
         .CellType(lRow, "IDPROD") = igCellCombo
         If .Combos("IDPROD").ListCount = 1 Then
            .CellValue(lRow, "IDPROD") = .Combos("IDPROD").ItemValue(0)
         End If
   
         .CellType(lRow, "QTDCOMPRA") = igCellText
         .CellValue(lRow, "QTDCOMPRA") = 0
   
         .CellType(lRow, "UNIDCONTROLE") = igCellText
         .CellValue(lRow, "UNIDCONTROLE") = ""
         If Val(.CellValue(lRow, "IDPROD")) <> 0 Then
            Call mvarMe_GrdPRODUTOAfterCommitEdit(lRow, .ColIndex("IDPROD"))
         End If
         
         For i = 1 To .ColCount
            If GetTag(.ColTag(i), "EDIT", True) = False Then
               .CellForeColor(lRow, i) = vbGrayText
            End If
         Next
      End If
      .AddRow
     
      For i = 1 To .ColCount
         .CellType(.RowCount, i) = igCellText
      Next
      For i = 1 To .ColCount
         If .ColVisible(i) Then
            .CellValue(.RowCount, i) = "Clique para incluir nova linha."
            .CellForeColor(.RowCount, i) = vbGrayText
            .CellTextFlags(.RowCount, i) = igTextNoClip Or igTextCenter
            Exit For
         End If
      Next
      
      
      .RowMode = (.RowCount = 1)
      If lRow <> 0 Then
         .SetCurCell .RowCount, 1
         .SetCurCell lRow, nColDefault
      End If
      If .Visible Then
         .SetFocus
      End If
      .Redraw = True
   End With
End Sub
Private Sub MontaGrdVenda()
'   Dim TbProd     As Object 'TB_SPRODUTO
'   Dim Sql        As String
'   Dim sAux       As String
   Dim i          As Integer
'
'   Set TbProd = CriarObjeto("BANCO_3R.TB_SPRODUTO")
'   Set TbProd.XDb = mvarSys.XDb
   
   On Error Resume Next
   With mvarMe.GrdVenda
      .ShowControlsInAllCells = True
      .Header.DragCols = False
      .SilentValidation = True
      .RowMode = True
      
      '.Header.ImageList = mvarMe.IlsIcons
      '.ImageList = mvarMe.IlsIcons
            
      
      With .AddCol(sKey:="ITEM", sHeader:="#", lWidth:=25, bVisible:=True)
         .eType = igCellText
         .sCtrlKey = "ITEM"
      End With
      
      
      With .AddCol(sKey:="IDVENDA", sHeader:="Venda", lWidth:=70, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextCenter
         .sCtrlKey = "IDVENDA"
      End With
      
      With .AddCol(sKey:="DTVENDA", sHeader:="Data", lWidth:=70, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextCenter
         .sCtrlKey = "DTVENDA"
      End With
            
      With .AddCol(sKey:="VLITENS", sHeader:="Venda", lWidth:=70, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight
         .sCtrlKey = "VLITENS"
      End With
      
      With .AddCol(sKey:="VLPGTO", sHeader:="Pagamento", lWidth:=70, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight
         .sCtrlKey = "VLPGTO"
      End With
      
      With .AddCol(sKey:="IDMOEDA", sHeader:="Moeda", lWidth:=70, bVisible:=False)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight
         .sCtrlKey = "IDMOEDA"
      End With
      
      For i = 1 To .ColCount
         .ColTag(i) = ""
      Next
      
      .ColTag("ITEM") = SetTag(.ColTag("ITEM"), "EDIT", False)
      .ColTag("IDVENDA") = SetTag(.ColTag("IDVENDA"), "EDIT", False)
      .ColTag("DTVENDA") = SetTag(.ColTag("DTVENDA"), "EDIT", False)
      .ColTag("VLITENS") = SetTag(.ColTag("VLITENS"), "EDIT", False)
      .ColTag("VLPGTO") = SetTag(.ColTag("VLPGTO"), "EDIT", False)
      .ColTag("IDMOEDA") = SetTag(.ColTag("IDMOEDA"), "EDIT", False)
            
      .ColTag("ITEM") = SetTag(.ColTag("ITEM"), "EDITOPT", igTextEditNumberOnly)

      IncluiGrdLinhaVenda

   End With
End Sub
Private Sub IncluiGrdLinhaVenda()
   Dim lRow As Long
   Dim i As Integer
   Dim nColDefault As Integer
         
   nColDefault = 2
   With mvarMe.GrdVenda
      'If .RowCount - 1 >= .Combos("IDVENDA").ListCount Then Exit Sub
   
      .Header.Font.Size = 9
      .Header.Font.Name = mvarMe.Font.Name
      .Font.Size = .Header.Font.Size
      .Font.Name = .Header.Font.Name
      
      lRow = .RowCount
      If lRow > 0 Then .RemoveRow lRow
      'If lRow > 1 Then .CellForeColor(lRow - 1, "IDVENDA") = vbBlack
      
      If lRow <> 0 Then
         .AddRow
         lRow = .RowCount
         .RowHeight(.RowCount) = 20

         .CellType(lRow, "ITEM") = igCellText
         .CellValue(lRow, "ITEM") = lRow

         .CellType(lRow, "IDVENDA") = igCellText
         .CellValue(lRow, "IDVENDA") = ""

         .CellType(lRow, "DTVENDA") = igCellText
         .CellValue(lRow, "DTVENDA") = ""

         .CellType(lRow, "VLITENS") = igCellText
         .CellValue(lRow, "VLITENS") = ""

         .CellType(lRow, "VLPGTO") = igCellText
         .CellValue(lRow, "VLPGTO") = ""

'         If Val(.CellValue(lRow, "IDVENDA")) <> 0 Then
'            Call mvarMe_GrdVendaAfterCommitEdit(lRow, .ColIndex("IDVENDA"))
'         End If

         For i = 1 To .ColCount
            If GetTag(.ColTag(i), "EDIT", True) = False Then
               .CellForeColor(lRow, i) = vbGrayText
            End If
         Next
      End If
      
      .AddRow
     
      For i = 1 To .ColCount
         .CellType(.RowCount, i) = igCellText
      Next
      For i = 1 To .ColCount
         If .ColVisible(i) Then
            .CellValue(.RowCount, i) = "Clique para incluir nova linha."
            .CellForeColor(.RowCount, i) = vbGrayText
            .CellTextFlags(.RowCount, i) = igTextNoClip Or igTextCenter
            Exit For
         End If
      Next
      
      
      '.RowMode = (.RowCount = 1)
      'If lRow <> 0 Then
      '   .SetCurCell .RowCount, 1
      '   .SetCurCell lRow, nColDefault
      'End If
      mvarMe.CmdEditar.Enabled = (.RowCount > 1)
      If .Visible Then
         .SetFocus
      End If
      .Redraw = True
   End With
End Sub
Private Function GetFOTOTIPOText(pFOTOTIPO As Integer, Optional pPele) As String
   Dim sAux As String
   Dim TbFoto As Object 'TB_OFOTOTIPO
   
   Select Case pFOTOTIPO
      Case 1: sAux = "I"
      Case 2: sAux = "II"
      Case 3: sAux = "III"
      Case 4: sAux = "IV"
      Case 5: sAux = "V"
      Case 6: sAux = "VI"
   End Select
   If IsMissing(pPele) Then
      Set TbFoto = CriarObjeto("BANCO_3R.TB_OFOTOTIPO", False)
      Set TbFoto.xDb = mvarSys.xDb
      If TbFoto.Pesquisar(Ch_FOTOTIPO:=pFOTOTIPO) Then
         GetFOTOTIPOText = sAux & " - " & TbFoto.PELE
      Else
         GetFOTOTIPOText = sAux
      End If
   Else
      GetFOTOTIPOText = sAux & " - " & pPele
   End If
   Set TbFoto = Nothing
End Function
Private Sub mvarMe_TxtHHFIMLostFocus()
   Dim bNull As Boolean
   
   bNull = IsNull(mvarMe.CmbHHFIM.Value)
   With mvarMe.TxtHHFIM
      If Not IsDate(FormatarHora(.Text)) Then
         If mvarCalEvent Is Nothing Then
            mvarMe.CmbHHFIM.Value = .CmbHHINI.Value
         Else
            mvarMe.CmbHHFIM.Value = mvarCalEvent.EndTime
         End If
         .Text = Format(mvarMe.CmbHHFIM.Value, "hh:mm")
      End If
      .Text = FormatarHora(.Text)
      If IsDate(.Text) Then
         mvarMe.CmbHHFIM.Value = .Text
         If bNull Then mvarMe.CmbHHFIM.Value = Null
      End If
   End With
   Call ExibeDuracao
End Sub
Private Sub ExibeDuracao()
   Dim sTempo As String
   Dim nTempo As Integer
   Dim nInt   As Integer
   
   On Error Resume Next
   If IsDate(mvarMe.TxtHHINI.Text) And IsDate(mvarMe.TxtHHFIM.Text) Then
      nTempo = DateDiff("n", mvarMe.TxtHHINI.Text, mvarMe.TxtHHFIM.Text)
      sTempo = ""
      
      mvarMe.LblDuracao.Caption = "Duração: "
      If nTempo > 0 Then
         nInt = Int(nTempo / 60)
         sTempo = StrZero(nInt, 2) & ":" & StrZero(nTempo Mod (IIf(nInt = 0, 1, nInt) * 60), 2)
      End If
      If sTempo <> "" Then
         mvarMe.LblDuracao.Caption = mvarMe.LblDuracao.Caption & sTempo & "h"
      End If
   End If
End Sub
Private Sub ExibeUltSessao()
   Dim Sql As String
   Dim MyRS As Object
   Dim sAux As String
   Dim nDias As Integer
   
   Sql = "Select Max(A.DTATEND) as [DTATEND], S.IDTPSERVICO "
   Sql = Sql & " From OATENDIMENTO  A "
   Sql = Sql & " Join OSESSAO S ON A.IDLOJA=S.IDLOJA And A.IDATENDIMENTO=S.IDATENDIMENTO "
   'Sql = Sql & " AND S.IDTPSERVICO<>1"
   Sql = Sql & " Where A.IDCLIENTE = " & SqlNum(mvarIDCLIENTE)
   Sql = Sql & " And A.IDLOJA =" & SqlNum(mvarIDLOJA)
   Sql = Sql & " And A.IDATENDIMENTO<>" & SqlNum(mvarIDATENDIMENTO)
   Sql = Sql & " And A.DTATEND <" & SqlDate(mvarMe.CmbDTATEND.Value)
   Sql = Sql & " Group By A.IDCLIENTE, S.IDTPSERVICO"
   Sql = Sql & " Order By S.IDTPSERVICO"
   
   mvarMe.LblUltSessao.Caption = ""
   With mvarSys.xDb
      If .AbreTabela(Sql, MyRS) Then
         If MyRS("IDTPSERVICO") = 1 Then
            nDias = DateDiff("d", MyRS("DTATEND"), mvarMe.CmbDTATEND.Value)
            sAux = "Avaliação em: "
            sAux = sAux & Format(MyRS("DTATEND"), "dd/mm/yy")
            sAux = sAux & " (" & CStr(nDias)
            sAux = sAux & " dias)"
            mvarMe.LblUltSessao.Caption = sAux
            nDias = 21
            MyRS.MoveNext
         End If
         If Not MyRS.EOF Then
            nDias = DateDiff("d", MyRS("DTATEND"), mvarMe.CmbDTATEND.Value)
            
            sAux = "Última Sessão: "
            sAux = sAux & Format(MyRS("DTATEND"), "dd/mm/yy")
            sAux = sAux & " (" & CStr(nDias)
            sAux = sAux & " dias)"
            mvarMe.LblUltSessao.Caption = sAux
         End If
      End If
      If nDias < 21 Or nDias > 39 Then
         mvarMe.LblUltSessao.ForeColor = 192
      Else
         mvarMe.LblUltSessao.ForeColor = &H404040
      End If
   End With
End Sub
Private Sub ExibeUltProd()
   Dim Sql As String
   Dim MyRS As Object
   Dim sAux As String
   Dim nDias As Integer
   Dim nQtd As Integer
   
   Sql = "Select V.IDCLIENTE, V.DTVENDA, I.IDPROD, I.QTDVENDA"
   Sql = Sql & " From CVENDA V"
   Sql = Sql & " Join CITENSVENDA I On V.IDLOJA=I.IDLOJA And V.IDVENDA=I.IDVENDA"
   Sql = Sql & " Join SPRODUTO P On P.IDPROD=I.IDPROD And P.ESERVICO=0"
   Sql = Sql & " Where V.IDCLIENTE = " & SqlNum(mvarIDCLIENTE)
   Sql = Sql & " And V.IDLOJA =" & SqlNum(mvarIDLOJA)
   Sql = Sql & " Order By V.DTVENDA"
   
   nQtd = 0
   nDias = 0
   mvarMe.LblProd.Caption = "Cliente NÃO comprou creme"
   With mvarSys.xDb
      If .AbreTabela(Sql, MyRS) Then
         While Not MyRS.EOF
            nQtd = nQtd + MyRS("QTDVENDA")
            MyRS.MoveNext
         Wend
         MyRS.MoveLast
         nDias = DateDiff("d", MyRS("DTVENDA"), mvarSys.xDb.Sysdate)
         
         sAux = "Cliente comprou " & nQtd & " creme" + IIf(nQtd <= 1, "", "s")
         sAux = sAux & " (" + Format(MyRS("DTVENDA"), "dd/mm/yy")
         sAux = sAux & " - "
         If nDias <= 30 Then
            sAux = sAux & CStr(nDias) & " dias)"
         Else
            sAux = sAux & CInt(nDias / 30) & " " & IIf(CInt(nDias / 30) <= 1, "mês)", "meses)")
         End If
         
         mvarMe.LblProd.Caption = sAux
      End If
      If (nDias / 30) > 6 Or nDias <= 0 Then
         mvarMe.LblProd.ForeColor = 192
      Else
         mvarMe.LblProd.ForeColor = &H404040
      End If
   End With
End Sub

Private Sub mvarMe_TxtHHINILostFocus()
   With mvarMe.TxtHHINI
      .Text = FormatarHora(.Text)
      If IsDate(.Text) Then
         mvarMe.CmbHHINI.Value = .Text
      End If
   End With
   Call ExibeDuracao
End Sub
Private Sub mvarMe_TxtIDCLIENTELostFocus()
   If xVal(mvarMe.TxtIDCLIENTE.Text) = 0 Then
      mvarMe.TxtIDCLIENTE.Text = GetTag(mvarMe.TxtNOME, "IDCLIENTE")
   ElseIf xVal(mvarMe.TxtIDCLIENTE.Text) <> xVal(GetTag(mvarMe.TxtNOME, "IDCLIENTE")) Then
      Call PopulaCliente(xVal(mvarMe.TxtIDCLIENTE.Text))
   End If
End Sub
'Private Sub mvarMe_TxtTEL1LostFocus()
'   Dim NG_Cal As NG_Calendario
'   Dim sTag As String
'
'   Set NG_Cal = New NG_Calendario
'   With NG_Cal
'      .IDLOJA = mvarIDLOJA
'      Set .Sys = mvarSys
'      sTag = .MProcuraCliente(mvarMe.TxtTEL1.Text, mvarMe.TxtNOME.Text, mvarIDCLIENTE)
'      If sTag <> "" Then
'         mvarMe.TxtNOME.Text = GetTag(sTag, "NOME", "")
'         mvarMe.TxtTEL1.Text = GetTag(sTag, "TEL1", "")
'         mvarMe.TxtFOTOTIPO.Text = GetFOTOTIPOText(GetTag(sTag, "FOTOTIPO", ""))
'         If mvarMe.CmbIDFUNCIONARIO.Enabled Then
'            mvarMe.CmbIDFUNCIONARIO.SetFocus
'         End If
'      End If
'   End With
'End Sub
Private Sub mvarMe_Unload(Cancel As Integer)
'   On Error Resume Next
'   Call RefreshEvent(mvarSys, mvarCalControl, mvarCalEvent, False)
End Sub
Private Sub AdicionarTratamento()
   Dim i As Integer
   Dim MySel As Collection
   Dim PreSel As Collection
   Dim nChave As Collection
   
   Set PreSel = New Collection
   With mvarMe.GrdSESSAO
      For i = 1 To .RowCount - 1
         If Not ExisteItem(PreSel, "k" & .CellValue(i, "IDTPTRATAMENTO")) Then
            PreSel.Add .CellValue(i, "IDTPTRATAMENTO"), "k" & .CellValue(i, "IDTPTRATAMENTO")
         End If
      Next
      Set MySel = SelecionarTratamento(PreSel)
      If Not MySel Is Nothing Then
         .Clear
         Call IncluiGrdLinhaSessao
         If .RowCount > 1 Then
            .RemoveRow 1
         End If
         i = 0
         For Each nChave In MySel
            Call IncluiGrdLinhaSessao
            i = i + 1
            .CellValue(i, "IDSERVICOEVT") = i
            If i = 1 Then
               .CellValue(i, "IDTPSERVICO") = 2 'IIf(mvarMe.ChkFLGAVALIACAO.Value, 1, 2)
            Else
               .CellValue(i, "IDTPSERVICO") = .CellValue(i - 1, "IDTPSERVICO")
            End If
            .CellValue(i, "IDTPTRATAMENTO") = xVal(nChave("IDTPTRATAMENTO"))
            .CellValue(i, "IDAREA") = 0
         Next
      End If
   End With
End Sub
Private Function SelecionarTratamento(Optional pPreSel As Collection) As Collection
   Dim MyAdd As Object
   Dim Sql As String
   
   Sql = "Select '' [CHK], IDTPTRATAMENTO, DSCTRATAMENTO"
   Sql = Sql + " From OTPTRATAMENTO"
   Sql = Sql + " Where ATIVO =1"
   Sql = Sql + " Order By DSCTRATAMENTO"
   
   Set MyAdd = CriarObjeto("XActive.DsAdd", False)
   If Not MyAdd Is Nothing Then
      With MyAdd
         .Limpar
         Set .xDb = mvarSys.xDb
         .Caption = "Tratamentos"
         .Chaves.Add "IDTPTRATAMENTO", "IDTPTRATAMENTO"
         
         With .TITULO
            '0-eAlin.Esquerda 1-eAlin.Direita 2-eAlin.Centro)
            Call .Add("CHK", "", 3, 0)
            Call .Add("IDTPTRATAMENTO", "Id.", 0, 0)
            Call .Add("DSCTRATAMENTO", "Descrição", 50, 0)
         End With
         Set .Preselecao = pPreSel
         .CheckBox = True
         .Query = Sql
   
         .Show
         Set SelecionarTratamento = .Selecao
      End With
   End If
End Function

