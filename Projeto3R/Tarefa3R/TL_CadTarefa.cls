VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CadTarefa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe    As FrmCadTarefa
Attribute mvarMe.VB_VarHelpID = -1

Private mvarsys            As Object
Private mvarIDLOJA         As Integer
Private mvarIsDirt         As Boolean
Private mvarIDTAREFA       As Long
Private mvarIDCLIENTE      As Long
Private mvarIDEVENTO       As Long
Private mvarTLCal          As Object
Public Property Set TLCal(ByVal vData As Object)
   Set mvarTLCal = vData
   mvarMe.CmdEvento.Enabled = Not (mvarTLCal Is Nothing)
End Property
Public Property Get TLCal() As Object
   Set TLCal = mvarTLCal
End Property
Public Property Get IsDirt() As Boolean
   IsDirt = mvarIsDirt
End Property
Public Property Let IDCLIENTE(ByVal vData As Long)
    mvarIDCLIENTE = vData
End Property
Public Property Get IDCLIENTE() As Long
   IDCLIENTE = mvarIDCLIENTE
End Property
Public Property Let IDEVENTO(ByVal vData As Long)
    mvarIDEVENTO = vData
End Property
Public Property Get IDEVENTO() As Long
   IDEVENTO = mvarIDEVENTO
End Property
Public Property Let IDTAREFA(ByVal vData As Long)
    mvarIDTAREFA = vData
End Property
Public Property Get IDTAREFA() As Long
   IDTAREFA = mvarIDTAREFA
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
    mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarsys = vData
   
   If mvarIDLOJA = 0 Then mvarIDLOJA = mvarsys.Propriedades("IDLOJA")
End Property
Public Property Get Sys() As Object
    Set Sys = mvarsys
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmCadTarefa
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
End Sub
Public Sub Show(Optional pMode As FormShowConstants = vbModal)
   mvarMe.Show pMode
End Sub
Private Sub mvarMe_Activate()
   'Call PopulaTela
   
   If GetTag(mvarMe, "1VEZ", "1") = "1" Then
      Call SetTag(mvarMe, "1VEZ", "0")
      If mvarMe.CmbTPTAREFA.Visible And mvarMe.CmbTPTAREFA.Enabled And mvarMe.CmbTPTAREFA.Text = "" Then
         mvarMe.CmbTPTAREFA.SetFocus
      ElseIf mvarMe.TxtTITULO.Enabled Then
         mvarMe.TxtTITULO.SetFocus
      ElseIf mvarMe.TxtDSCTAREFA.Enabled Then
         mvarMe.TxtDSCTAREFA.SetFocus
         mvarMe.TxtDSCTAREFA.SelStart = Len(mvarMe.TxtDSCTAREFA.Text)
      End If
   End If
End Sub

Private Sub mvarMe_CmbTPTAREFAClick()
   mvarMe.TxtTITULO.Enabled = (mvarMe.CmbTPTAREFA.ItemData(mvarMe.CmbTPTAREFA.ListIndex) = 0)
   If Not mvarMe.TxtTITULO.Enabled Then
      Call MontaTitulo
   Else
      If GetTag(mvarMe.TxtTITULO, "TITULO", "") <> "" Then
         mvarMe.TxtTITULO.Text = GetTag(mvarMe.TxtTITULO, "TITULO", "")
      End If
   End If
   If mvarMe.CmdEMAIL.Caption = "" Then
      mvarMe.CmdeMail2.Enabled = False
   Else
      mvarMe.CmdeMail2.Enabled = Between(mvarMe.CmbTPTAREFA.ItemData(mvarMe.CmbTPTAREFA.ListIndex), eTpTarefa.TarBoasVindas, eTpTarefa.TarNaoMarcado)
   End If
End Sub

Private Sub mvarMe_CmbTPTAREFALostFocus()
   If LocalizarCombo(mvarMe.CmbTPTAREFA, mvarMe.CmbTPTAREFA.Text) < 0 Then
      mvarMe.CmbTPTAREFA.ListIndex = 0
   End If
End Sub
Private Sub mvarMe_CmdCancelClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdEMAIL2Click()
   Dim MyObj As Object
   Dim bResult As Boolean
   
   If Between(mvarMe.CmbTPTAREFA.ItemData(mvarMe.CmbTPTAREFA.ListIndex), eTpTarefa.TarBoasVindas, eTpTarefa.TarNaoMarcado) Then
      bResult = ExibireMail(mvarIDTAREFA)
      If bResult Then
         Call LocalizarCombo(mvarMe.CmbSITTAREFA, 10, True, True)
         mvarMe.TxtDSCTAREFA.Text = mvarMe.TxtDSCTAREFA.Text & vbNewLine & "[" & Format(mvarsys.xDb.SysDate(3), "dd/mm/yy hh:mm") & "] - e-Mail enviado. "
         mvarMe.TxtDSCTAREFA.SetFocus
         mvarMe.TxtDSCTAREFA.SelStart = Len(mvarMe.TxtDSCTAREFA.Text)
      Else
         'If vbYes = ExibirPergunta("Deseja abrir seu programa de email padrão?", "Correio", False) Then
         '   Call ExecuteLink("mailto:" & mvarMe.CmdEMAIL.Caption)
         'End If
      End If
   Else
      Call ExecuteLink("mailto:" & mvarMe.CmdEMAIL.Caption)
   End If
End Sub

Private Sub mvarMe_CmdEMAILClick()
   Call ExecuteLink("mailto:" & mvarMe.CmdEMAIL.Caption)
End Sub
Public Function ExibireMail(pIDTAREFA As Long) As Boolean
   Dim sTit       As String
   Dim sPath      As String
   Dim sNMTarefa  As String
   Dim sPathTmp   As String
   Dim sAttach    As String
   Dim sFile      As String
   Dim sHtml      As String
   Dim pAttach    As String
   Dim pFile      As String
   Dim Sql        As String
   Dim Rs         As Object
   Dim MyObj      As Object
   Dim cAux       As Collection
   Dim bResult    As Boolean
   Dim i          As Integer
   Dim nIDEVENTO  As Long
   Dim nTPTAREFA  As Long
   
   '********
   '* Recuperar tipo de tarefa/email.
   Sql = "Select T.IDEVENTO, T.IDTPTAREFA, TT.EMAILKEY, TT.EMAILTIT " & vbNewLine
   Sql = Sql & " From OTAREFAEVT T" & vbNewLine
   Sql = Sql & " Join OTIPOTAREFA TT On T.IDLOJA=TT.IDLOJA And T.IDTPTAREFA=TT.IDTPTAREFA" & vbNewLine
   Sql = Sql & " Where T.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And T.IDTAREFA=" & SqlNum(pIDTAREFA) & vbNewLine
   If mvarsys.xDb.AbreTabela(Sql, Rs) Then
      nTPTAREFA = xVal(Rs("IDTPTAREFA") & "")
      nIDEVENTO = xVal(Rs("IDEVENTO") & "")
      sTit = Rs("EMAILTIT")
      sNMTarefa = Rs("EMAILKEY")
   Else
      Call ExibirAviso("Erro ao selecionar evento.")
      ExibireMail = False
      Exit Function
   End If
   
   '********
   '* Definir Pastas de Trabalho
   If InStr(UCase(mvarsys.ExePath), UCase("C:\Sistemas")) <> 0 Then
      sPath = Environ("programfiles") & "\ClasseA\Projeto3R\"
   Else
      sPath = mvarsys.ExePath
   End If
   sPath = sPath & "HTM\"
   sPathTmp = Environ("TEMP") & "\" & mvarsys.CODSIS & "\"
   
   '********
   '* Definir Arquivo de Trabalho
   sAttach = sNMTarefa & ".gif"
   sFile = sNMTarefa & ".htm"
   
   '********
   '* Limpar Pasta de Trabalho
   pAttach = sPathTmp & sAttach
   pFile = sPathTmp & sFile
   If ExisteArquivo(pAttach) Then Call ExcluirArquivo(pAttach, False)
   If ExisteArquivo(pFile) Then Call ExcluirArquivo(pFile, False)
   
   '********
   '* Recuperar HTML
   If ExisteArquivo(sPath & sFile) Then
      Call CopiarArquivo(sPath & sFile, pFile)
   Else
      Call ExtractResData(UCase(sNMTarefa), "HTML", pFile)
   End If
   '********
   '* Recuperar Arquivo Anexo
   If ExisteArquivo(sPath & sAttach) Then
      Call CopiarArquivo(sPath & sAttach, pAttach)
   Else
      Call ExtractResData(UCase(sNMTarefa), "GIF", pAttach)
   End If
   
   If ExisteArquivo(pFile) Then sHtml = ReadTextFile(pFile)
   
   If Trim(sHtml) <> "" Then
      Dim sNome As String, sData As String, sDIASEMANA As String, sHORA As String
      Dim sUNIDADE As String, sENDERECO As String, sTELEFONE As String
      Dim sDestino As String
           
      Sql = "Select E.STARTDATETIME, C.NOME, C.SEXO, C.EMAIL, L.NOME [NMLOJA]"
      Sql = Sql & ", L.ENDERECO, L.TELEFONE1"
      Sql = Sql & " From OEVENTOAGENDA E"
      Sql = Sql & " JOIN OCLIENTE C On E.IDLOJA=C.IDLOJA AND E.IDCLIENTE=C.IDCLIENTE"
      Sql = Sql & " JOIN OLOJA L On E.IDLOJA=L.IDLOJA"
      Sql = Sql & " Where E.IDLOJA=" & mvarsys.IDLOJA
      Sql = Sql & " And E.IDEVENTO=" & nIDEVENTO
      If mvarsys.xDb.AbreTabela(Sql, Rs) Then
         sNome = Rs("NOME") & ""
         Set cAux = GetWords(sNome)
         sNome = "Sr(a)"
         For i = 1 To cAux.Count
            sNome = sNome & " " & cAux(i)
            If i <> 1 Then
              If Len(cAux(i)) > 3 Then
                 Exit For
              End If
            End If
         Next
         sData = Format(Rs("STARTDATETIME"), "DD/MM/YY")
         sDIASEMANA = StrConv(WeekdayName(Weekday(sData)), vbProperCase)
         sHORA = Format(Rs("STARTDATETIME"), "HH:MM")
         sUNIDADE = mvarsys.Propriedades("FromDisplayName") 'Rs("NMLOJA")
         sENDERECO = Rs("ENDERECO")
         sTELEFONE = Rs("TELEFONE1")
                     
         sDestino = Rs("EMAIL")
      End If
      '**************************
      '***** POPULAR HTML *******
      '**************************
      sHtml = Replace(sHtml, "@@NOME", sNome)
      sHtml = Replace(sHtml, "@@DATA", sData)
      sHtml = Replace(sHtml, "@@DIASEMANA", sDIASEMANA)
      sHtml = Replace(sHtml, "@@HORA", sHORA)

      sHtml = Replace(sHtml, "@@UNIDADE", sUNIDADE)
      sHtml = Replace(sHtml, "@@ENDERECO", sENDERECO)
      sHtml = Replace(sHtml, "@@TELEFONE", sTELEFONE)
   End If
   
   i = AbrirTxt(pFile)
   Call EscreverTxt(i, sHtml)
   Call FecharTxt(i)
   
   If Trim(sHtml) <> "" Then
      Set MyObj = CriarObjeto("Utilitario3R.NG_UTILITARIO", False)
      Set MyObj.Sys = mvarsys
      bResult = MyObj.F_Enviar_eMail(pFile, sTit, sHtml, sAttach, sNome, sDestino)
   End If
   Set MyObj = Nothing
   ExibireMail = bResult
End Function
Private Sub mvarMe_CmdEventoClick()
   Dim TlEvento As Object
   Dim TbEvento As Object
   
   If mvarIDEVENTO = 0 Then Exit Sub
   'If TypeName(mvarTLCal) = "" Then
      mvarTLCal.TlEvento.IDEVENTO = mvarIDEVENTO
      mvarTLCal.TlEvento.Show
      
      If mvarMe.CmbTPTAREFA.ItemData(mvarMe.CmbTPTAREFA.ListIndex) = eTpTarefa.TarConfimaAge Then
         Set TbEvento = CriarObjeto("Banco_3R.TB_OEVENTOAGENDA", False)
         Set TbEvento.xDb = mvarsys.xDb
         If TbEvento.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEVENTO:=mvarTLCal.TlEvento.IDEVENTO) Then
            If TbEvento.FLGCONFIRMADO = 1 Then
               Call LocalizarCombo(mvarMe.CmbSITTAREFA, 10, True, True)
            End If
         End If
      End If
      
   'ElseIf TypeName(mvarTLCal) = "" Then
   '   mvarTLCal.TlEvento.IDEVENTO = mvarIDEVENTO
   '   mvarTLCal.TlEvento.Show
   'End If
End Sub
Private Sub mvarMe_CmdExcluirClick()
   Dim sMsg As String
   Dim TbTarefa As Object
   
   sMsg = "Esta operação irá excluir a tarefa Nº'" & mvarMe.TxtIDTAREFA.Text & "'"
   sMsg = sMsg & vbNewLine & vbNewLine
   sMsg = sMsg & "Deseja continuar?"
   If ExibirPergunta(sMsg, , False) = vbYes Then
      If InputBoxPassword("Informe a senha.") = mvarsys.Propriedades("SENHAMESTRE") Then
         Call PopulaClasse(TbTarefa)
         Call TbTarefa.Salvar
      Else
         Call ExibirAviso("Esta senha não permite a operação.")
      End If
   End If
End Sub
Private Sub mvarMe_CmdIDCLIENTEClick()
   Call ShowCliente(mvarsys, mvarIDLOJA, mvarIDCLIENTE)
End Sub
Private Sub mvarMe_CmdLovCliClick()
   Dim MyCliente As Object
   
   Set MyCliente = CriarObjeto("Contato3R.TL_NovoLov", False)
   With MyCliente
      .Limpar
      Set .Sys = mvarsys
      .Filtro = Trim(mvarMe.TxtNOME.Text)
      Call .ShowClientes
      If .IdSelecionado <> 0 Then
         Call PopulaCliente(.IdSelecionado)
      End If
   End With
   Set MyCliente = Nothing
End Sub
Private Function ValidaCampos() As Boolean
   Dim i As Long
   
'   If mvarIDCLIENTE = 0 Then
'      Call ExibirInformacao("Informe o cliente da tarefa.")
'      If mvarMe.CmdLovCli.Enabled Then
'         Call mvarMe_CmdLovCliClick
'      End If
'      Exit Function
'   End If
   If mvarMe.CmbTPTAREFA.Text = "" Then
      Call ExibirInformacao("Informe o tipo da tarefa.")
      If mvarMe.CmbTPTAREFA.Enabled Then
         mvarMe.CmbTPTAREFA.SetFocus
      End If
      Exit Function
   End If
   
   If Trim(mvarMe.TxtTITULO.Text) = "" Then
      Call ExibirInformacao("Informe o título da tarefa.")
      If mvarMe.TxtTITULO.Enabled Then
         mvarMe.TxtTITULO.SetFocus
      End If
      Exit Function
   End If
   If Not IsDate(mvarMe.CmbDTTAREFA.Value) Then
      Call ExibirInformacao("Informe a data da tarefa.")
      If mvarMe.CmbDTTAREFA.Enabled Then
         mvarMe.CmbDTTAREFA.SetFocus
      End If
      Exit Function
   End If
   If Trim(mvarMe.TxtDSCTAREFA.Text) = "" Then
      Call ExibirInformacao("Informe os detalhes da tarefa.")
      If mvarMe.TxtDSCTAREFA.Enabled Then
         mvarMe.TxtDSCTAREFA.SetFocus
      End If
      Exit Function
   End If
   

   ValidaCampos = True
End Function
Private Function SalvarTarefa() As Boolean
   Dim bOk     As Boolean
   Dim bIsDirt As Boolean
   Dim TbTarefa As Object
   
   Screen.MousePointer = vbHourglass
   If ValidaCampos() Then
      Call PopulaClasse(TbTarefa)
      
      bIsDirt = TbTarefa.IsDirt
      bOk = True
      If bIsDirt Then
         bOk = TbTarefa.Salvar()
         If bOk Then
            mvarIDTAREFA = TbTarefa.IDTAREFA
            
            If mvarMe.CmbTPTAREFA.ItemData(mvarMe.CmbTPTAREFA.ListIndex) = eTpTarefa.TarConfimaAge Then
               Call ConfirmaAgenda(mvarsys, mvarIDLOJA, mvarIDTAREFA, mvarTLCal)
            End If
            
         End If
         mvarIsDirt = mvarIsDirt Or bOk
         Call ExibirResultado(mvarsys, bOk, 2)
      End If
      SalvarTarefa = bOk
   End If
   Screen.MousePointer = vbDefault
End Function
Public Sub ConfirmaAgenda(pSys As Object, pIDLOJA As Integer, pIDTAREFA As Long, pTLCal As Object)
   Dim TbEvento As Object
   Dim TbTarefa As Object
   
   Dim nIDEVENTO As Long
   Dim nSITTAREFA As Integer
   
   Set TbTarefa = CriarObjeto("BANCO_3R.TB_OTAREFAEVT", False)
   Set TbTarefa.xDb = mvarsys.xDb
   If TbTarefa.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDTAREFA:=mvarIDTAREFA) Then
      nIDEVENTO = TbTarefa.IDEVENTO
      nSITTAREFA = TbTarefa.SITTAREFA
   Else
      Exit Sub
   End If
   Set TbTarefa = Nothing

   
   Set TbEvento = CriarObjeto("Banco_3R.TB_OEVENTOAGENDA", False)
   Set TbEvento.xDb = mvarsys.xDb
   If TbEvento.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEVENTO:=nIDEVENTO) Then
      If nSITTAREFA = 10 Then
         TbEvento.FLGCONFIRMADO = 1
      Else
         TbEvento.FLGCONFIRMADO = 0
      End If
      If TbEvento.IsDirt Then
         If TbEvento.Alterar Then
            If Not pTLCal Is Nothing Then
               If Not pTLCal.TlEvento Is Nothing Then
                  pTLCal.TlEvento.IDEVENTO = nIDEVENTO
                  If Not pTLCal.TlEvento.CalEvent Is Nothing Then
                     pTLCal.TlEvento.CalEvent.CustomProperties("FLGCONFIRMADO") = TbEvento.FLGCONFIRMADO
                     Call RefreshEvent(pSys, pTLCal.TlEvento.Calcontrol, pTLCal.TlEvento.CalEvent)
                  End If
               End If
            End If
         End If
      End If
   End If
   Set TbEvento = Nothing
End Sub
Private Sub PopulaClasse(pTBTarefa As Object, Optional bDelete As Boolean = False)
   If pTBTarefa Is Nothing Then
      Set pTBTarefa = CriarObjeto("BANCO_3R.TB_OTAREFAEVT", False)
      Call pTBTarefa.Limpar
      Set pTBTarefa.xDb = mvarsys.xDb
      Call pTBTarefa.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDTAREFA:=mvarIDTAREFA)
   End If
   With pTBTarefa
      .IDLOJA = mvarIDLOJA
      .IDEVENTO = mvarMe.CmbIDEVENTO.ItemData(mvarMe.CmbIDEVENTO.ListIndex)
      .IDCLIENTE = mvarIDCLIENTE
      .IDTPTAREFA = mvarMe.CmbTPTAREFA.ItemData(mvarMe.CmbTPTAREFA.ListIndex)
      .DSCTAREFA = mvarMe.TxtDSCTAREFA
      .TITULO = mvarMe.TxtTITULO
      .DTTAREFA = mvarMe.CmbDTTAREFA.Value
      .SITTAREFA = mvarMe.CmbSITTAREFA.ItemData(mvarMe.CmbSITTAREFA.ListIndex)
      .Prioridade = mvarMe.CmbPrioridade.ItemData(mvarMe.CmbPrioridade.ListIndex)
      '.FLGLEMBRETE
      '.DLEMBRETE
      .FLGDELETE = IIf(bDelete, 1, 0)
      '.FLGAUTO
   End With
End Sub
Public Sub ExcluirTarefa()

End Sub

Private Sub mvarMe_CmdOkClick()
   If SalvarTarefa Then
      Unload mvarMe
   End If
End Sub

Private Sub mvarMe_Load()
   mvarMe.Icon = mvarsys.MDI.Icon
   '* Monta Tela
   Call MontaTela
   Call PopulaTela
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub MontaTela()
   Dim TbTipo As Object
   With mvarMe
      .CmdExcluir.Enabled = False
      .CmdEvento.Enabled = False
      mvarMe.CmdEMAIL.Enabled = True
      mvarMe.CmdeMail2.Enabled = False
      .TxtTEL1.MaxLength = mvarsys.Propriedades("MaxLenTel")
      
      .CmbTPTAREFA.Clear
      .CmbTPTAREFA.AddItem " Outras Tarefas"
      .CmbTPTAREFA.ItemData(mvarMe.CmbTPTAREFA.NewIndex) = 0
      
      Set TbTipo = CriarObjeto("BANCO_3R.TB_OTIPOTAREFA", False)
      With TbTipo
         .Limpar
         Set .xDb = mvarsys.xDb
         If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="ATIVO=1", Ch_OrderBy:="IDTPTAREFA") Then
            While Not .Rs.EOF
               .Popula
               mvarMe.CmbTPTAREFA.AddItem .DSCTPTAREFA
               mvarMe.CmbTPTAREFA.ItemData(mvarMe.CmbTPTAREFA.NewIndex) = .IDTPTAREFA
               .Rs.MoveNext
            Wend
         End If
      End With
      .CmbTPTAREFA.ListIndex = 0
      
      .CmbPrioridade.Clear
      .CmbPrioridade.AddItem "Alta", 0
      .CmbPrioridade.AddItem "Normal", 1
      .CmbPrioridade.ItemData(0) = 1
      .CmbPrioridade.ItemData(1) = 2
      .CmbPrioridade.ListIndex = 1
      
      .CmbSITTAREFA.Clear
      .CmbSITTAREFA.AddItem "Não Iniciada", 0
      .CmbSITTAREFA.AddItem "Em Andamento", 1
      .CmbSITTAREFA.AddItem "Concluída", 2
      .CmbSITTAREFA.ItemData(0) = 0
      .CmbSITTAREFA.ItemData(1) = 5
      .CmbSITTAREFA.ItemData(2) = 10
      .CmbSITTAREFA.ListIndex = 0
      
      
   End With
End Sub
Private Sub PopulaTela()
   Dim Sql        As String
   Dim TbTarefa   As Object
   
   Dim TbEvento   As Object
   Dim bOk     As Boolean
   
   Call LimparTela
   
   Set TbTarefa = CriarObjeto("BANCO_3R.TB_OTAREFAEVT", False)
   With TbTarefa
      .Limpar
      Set .xDb = mvarsys.xDb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDTAREFA:=mvarIDTAREFA) Then
         Call SetTag(mvarMe.TxtTITULO, "TITULO", .TITULO)
         mvarIDEVENTO = .IDEVENTO
         
         mvarMe.Caption = "Tarefa Nº" & StrZero(.IDTAREFA, 6)
         mvarMe.TxtDSCTAREFA.Text = .DSCTAREFA
         mvarMe.TxtIDTAREFA.Text = .IDTAREFA
         mvarMe.TxtTITULO.Text = .TITULO
         mvarMe.TxtIDEVENTO.Text = .IDEVENTO
         mvarMe.CmbDTTAREFA.Value = .DTTAREFA
         Call LocalizarCombo(mvarMe.CmbPrioridade, .Prioridade, True, True)
         Call LocalizarCombo(mvarMe.CmbSITTAREFA, .SITTAREFA, True, True)
         Call LocalizarCombo(mvarMe.CmbTPTAREFA, .IDTPTAREFA, True, True)
         
         Call PopulaCliente(pIDCLIENTE:=.IDCLIENTE)
         Call LocalizarCombo(mvarMe.CmbIDEVENTO, .IDEVENTO, True, True)
                                    
         If mvarMe.CmbPrioridade.ListIndex = -1 Then
            Call LocalizarCombo(mvarMe.CmbPrioridade, 2, True, True)
         End If
         If mvarMe.CmbSITTAREFA.ListIndex = -1 Then
            Call LocalizarCombo(mvarMe.CmbSITTAREFA, 0, True, True)
         End If
         
         Call HabilitarCampos((.FLGAUTO = 1))
         If mvarMe.TxtDSCTAREFA.Visible Then
            mvarMe.TxtDSCTAREFA.SetFocus
            mvarMe.TxtDSCTAREFA.SelStart = Len(mvarMe.TxtDSCTAREFA.Text)
         End If
                  
         mvarMe.CmdeMail2.Enabled = False
         If Trim(mvarMe.CmdEMAIL.Caption) <> "" Then
            If Between(mvarMe.CmbTPTAREFA.ItemData(mvarMe.CmbTPTAREFA.ListIndex), eTpTarefa.TarBoasVindas, eTpTarefa.TarNaoMarcado) Then
               mvarMe.CmdeMail2.Enabled = True
            End If
         End If
      Else
         If mvarIDCLIENTE <> 0 Then
            Call PopulaCliente(pIDCLIENTE:=mvarIDCLIENTE)
            
            If mvarMe.CmbIDEVENTO.ListIndex = -1 Then
               mvarMe.CmbIDEVENTO.ListIndex = mvarMe.CmbIDEVENTO.ListCount - 1
            End If
            If mvarMe.CmbPrioridade.ListIndex = -1 Then
               Call LocalizarCombo(mvarMe.CmbPrioridade, 2, True, True)
            End If
            If mvarMe.CmbSITTAREFA.ListIndex = -1 Then
               Call LocalizarCombo(mvarMe.CmbSITTAREFA, 0, True, True)
            End If
            Call HabilitarCampos(False)
            mvarMe.CmdExcluir.Enabled = False
         End If
         If mvarMe.TxtTITULO.Visible And mvarMe.TxtTITULO.Enabled Then
            mvarMe.TxtTITULO.SetFocus
         End If
      End If
   End With
End Sub
Private Sub LimparTela()
   Dim n As Object
   
   On Error Resume Next
   mvarMe.Caption = "Tarefa"
   For Each n In mvarMe.Controls
      n.Text = ""
      n.ListIndex = -1
      n.Value = Nothing
   Next
   mvarMe.CmbPrioridade.ListIndex = 1
   mvarMe.CmbSITTAREFA.ListIndex = 0
   
   mvarMe.CmdEMAIL.Caption = ""
   mvarMe.CmbDTTAREFA.Value = mvarsys.xDb.SysDate()
End Sub
Private Sub PopulaCliente(Optional pIDCLIENTE As Long, Optional pTelefone As String = "")
   Dim TbCli   As Object 'TB_OCLIENTE
   Dim bAchou  As Boolean
   
   Set TbCli = CriarObjeto("BANCO_3R.TB_OCLIENTE", False)
   With TbCli
      .Limpar
      Set .xDb = mvarsys.xDb
      If IsMissing(pIDCLIENTE) Then
         bAchou = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="TEL1='" & pTelefone & "'")
         If Not bAchou Then bAchou = .Pesquisar(Ch_Where:="TEL2='" & pTelefone & "'")
      Else
         bAchou = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCLIENTE:=pIDCLIENTE)
      End If
      If bAchou Then
         mvarIDCLIENTE = .IDCLIENTE
         Call SetTag(mvarMe.TxtNOME, "IDCLIENTE", .IDCLIENTE)
         mvarMe.TxtNOME.Text = .NOME
         mvarMe.TxtTEL1.Text = IIf(.TEL1 = "", IIf(.TEL2 = "", .FAX, .TEL2), .TEL1)
         mvarMe.CmdEMAIL.Caption = .eMail
      Else
         mvarIDCLIENTE = 0
         Call SetTag(mvarMe.TxtNOME, "IDCLIENTE", "")
         mvarMe.TxtNOME.Text = ""
         mvarMe.TxtTEL1.Text = ""
      End If
      Call MontaEvento
   End With
   Set TbCli = Nothing

End Sub
Private Sub mvarMe_TxtDSCTAREFAKeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      If Right(mvarMe.TxtDSCTAREFA.Text, 1) = Chr(10) Then
         mvarMe.TxtDSCTAREFA.Text = mvarMe.TxtDSCTAREFA.Text & "[" & Format(mvarsys.xDb.SysDate(3), "dd/mm/yy hh:mm") & "] - "
         mvarMe.TxtDSCTAREFA.SelStart = Len(mvarMe.TxtDSCTAREFA.Text)
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub mvarMe_TxtNOMEChange()
   Call MontaTitulo
End Sub
Private Sub mvarMe_TxtTEL1Change()
   Call MontaTitulo
End Sub
Private Sub MontaTitulo()
   Dim sNome As String
   Dim sTel  As String
   Dim sAux  As String
   Dim i     As Integer
   
   With mvarMe
      If .CmbTPTAREFA.ItemData(.CmbTPTAREFA.ListIndex) <> 0 Then
         sNome = Trim(.TxtNOME.Text)
         sTel = Trim(.TxtTEL1.Text)
         
         For i = 1 To Len(sTel)
            If isAlfaNum(Mid(sTel, i, 1)) Then
               sAux = sAux & Mid(sTel, i, 1)
            End If
         Next
         sTel = sAux
         
         .TxtTITULO.Text = sNome & IIf(sTel = "", "", " - ") & sTel
      End If
   End With
End Sub
Private Sub MontaEvento()
   Dim Sql As String
   Dim MyRs As Object
      
   If mvarIDCLIENTE <> xVal(SetTag(mvarMe.TxtNOME, "IDCLIENTE", "0")) Then
      Sql = "Select *"
      Sql = Sql & " From OEVENTOAGENDA"
      Sql = Sql & " Where IDCLIENTE= " & mvarIDCLIENTE
      Sql = Sql & " And IDLOJA= " & mvarIDLOJA
      Sql = Sql & " Order By StartDateTime"
      mvarMe.CmbIDEVENTO.Clear
      If mvarsys.xDb.AbreTabela(Sql, MyRs) Then
         While Not MyRs.EOF
            mvarMe.CmbIDEVENTO.AddItem MyRs("STARTDATETIME")
            mvarMe.CmbIDEVENTO.ItemData(mvarMe.CmbIDEVENTO.NewIndex) = MyRs("IDEVENTO")
            MyRs.MoveNext
         Wend
      End If
   End If
End Sub
Private Sub HabilitarCampos(pAuto As Boolean)
   mvarMe.TxtTITULO.Enabled = Not pAuto
   mvarMe.CmbIDEVENTO.Enabled = Not pAuto
   mvarMe.CmbTPTAREFA.Enabled = Not pAuto
   mvarMe.CmdExcluir.Enabled = Not pAuto
End Sub

Private Sub mvarMe_TxtVLDESCChange()

End Sub
