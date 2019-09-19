VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_FecharDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim nTotalErros   As Integer
Dim nTotalAvisos  As Integer
Dim bExeFechamento As Boolean

Private Enum Bandeira
    Verde = 1
    Vermelha = 2
    Amarela = 3
    Ok = 4
    Erro = 5
    Aviso = 6
End Enum
Private WithEvents mvarMe  As FrmFecharDia
Attribute mvarMe.VB_VarHelpID = -1

Private mvarSys         As Object
Private mvarCalControl  As Object
Private mvarNGCal       As Object
Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Property Set CalControl(ByVal vData As Object)
    Set mvarCalControl = vData
End Property
Public Property Get CalControl() As Object
    Set CalControl = mvarCalControl
End Property
Public Property Set NGCal(ByVal vData As Object)
    Set mvarNGCal = vData
End Property
Public Property Get NGCal() As Object
    Set NGCal = mvarNGCal
End Property
Public Sub Show()
   '* Carregar Form
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmFecharDia
End Sub
Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call FecharDia(mvarMe.CmbDTFECHAR.Value)
      Call SetTag(mvarMe, "1VEZ", 0)
   End If
End Sub
Private Sub mvarMe_ChkExibirClick()
   Call ExibirItens
End Sub
Private Sub mvarMe_CmdFecharDia()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmbDTFECHARChange()
   Call LimparGrid
End Sub
Private Sub mvarMe_CmdFecharDiaClick()
   Call FecharDia(mvarMe.CmbDTFECHAR.Value)
End Sub
Private Sub mvarMe_CmdSairClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_GrdFecharDiaHyperlinkClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal HyperlinkIndex As Long)
   Dim sCod As String
   
   sCod = xVal(Mid(Item.Value, Item.Hyperlinks.Item(0).HyperTextBegin + 1, Item.Hyperlinks.Item(0).HyperTextLen))
   
   Select Case GetTag(Item, "TABELA", "")
      Case "OEVENTOAGENDA": Call ExibirEvento(sCod)
      Case "OATENDIMENTO": Call ExibirAtendimento(sCod)
      Case "CVENDA": Call ExibirVenda(sCod)
   End Select
   'Call FecharDia(mvarMe.CmbDTFECHAR.Value)
End Sub
Private Sub ExibirEvento(pCod As String)
   Dim TlEvento As Object
   Dim TlAtend As Object
   
   If mvarNGCal Is Nothing Then
      Set mvarNGCal = CriarObjeto("Calendario3R.NG_Calendario")
      Set mvarNGCal.Sys = mvarSys
   End If
   
   If mvarNGCal.ExisteAtendimento(Val(pCod)) <> 0 Then
      Set TlAtend = CriarObjeto("Clalendario3R.TL_Atendimento")
      With TlAtend
         Set .Sys = mvarSys
         Set .CalControl = mvarCalControl
         Set .CalEvent = mvarCalControl.DataProvider.GetEvent(pCod)
         .IDEVENTO = Val(pCod)
         .Show vbModal
      End With
   Else
      
      Set TlEvento = CriarObjeto("Calendario3R.TL_Evento")
      Set mvarNGCal.TlEvt = TlEvento
      Set TlEvento.CalControl = mvarCalControl
      Set TlEvento.CalEvent = mvarCalControl.DataProvider.GetEvent(pCod)
      Set TlEvento.Sys = mvarSys
      TlEvento.Show
   End If
      
End Sub
Private Sub ExibirAtendimento(pCod As String)
   Dim TBAtend As Object
   Dim TlAtend As Object
   
   Set TBAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
   Set TBAtend.xdb = mvarSys.xdb
   If TBAtend.Pesquisar(Ch_IDLOJA:=mvarSys.Propriedades("IDLOJA"), Ch_IDATENDIMENTO:=pCod) Then
      Set TlAtend = CriarObjeto("Calendario3R.TL_Atendimento")
      With TlAtend
         Set .Sys = mvarSys
         Set .CalControl = mvarCalControl
         Set .CalEvent = mvarCalControl.DataProvider.GetEvent(TBAtend.IDEVENTO)
         .IDEVENTO = TBAtend.IDEVENTO
         .Show vbModal
      End With
   End If
   
   Set TBAtend = Nothing
   Set TlAtend = Nothing
End Sub
Private Sub ExibirVenda(pCod As String)
   Dim TBVenda As Object
   Dim TBAtend As Object
   Dim TlVenda As Object
   
   Set TBVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
   Set TBVenda.xdb = mvarSys.xdb
   If TBVenda.Pesquisar(Ch_IDLOJA:=mvarSys.Propriedades("IDLOJA"), Ch_IDVENDA:=pCod) Then
      Set TlVenda = CriarObjeto("Calendario3R.TL_Venda")
      With TlVenda
         Set TBAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_VENDA")
         Set TBAtend.xdb = mvarSys.xdb
         If TBAtend.Pesquisar(Ch_IDLOJA:=mvarSys.Propriedades("IDLOJA"), Ch_IDVENDA:=pCod) Then
            .IDATENDIMENTO = TBAtend.IDATENDIMENTO
         End If
         
         Set .Sys = mvarSys
         Set .CalControl = mvarCalControl
         .IDLOJA = TBVenda.IDLOJA
         .IDCLIENTE = TBVenda.IDCLIENTE
         .IDVENDA = TBVenda.IDVENDA
         
         .Show vbModal
      End With
      Set TlVenda = Nothing
      Set TBAtend = Nothing
      Set TBVenda = Nothing
   End If
End Sub
Private Sub mvarMe_Load()
   Dim sDT  As String
      
   If Not mvarCalControl Is Nothing Then sDT = Format(mvarCalControl.DayView.Selection.Begin, "dd/mm/yyyy")
   If Not IsDate(sDT) Then sDT = mvarSys.xdb.Sysdate(1)
   mvarMe.CmbDTFECHAR.Value = sDT
End Sub
Private Sub FecharDia(pData As String)
   Dim sDate As String
   
   Screen.MousePointer = vbHourglass
   
   sDate = Format(pData, "dd/mm/yyyy")
   
   Call LimparGrid
   
   mvarMe.ProgressBar1.Scrolling = xtpProgressBarMarquee
   mvarMe.ProgressBar1.Enabled = True
   mvarMe.ProgressBar1.Visible = True
   DoEvents
   
   nTotalErros = 0
   nTotalAvisos = 0
   
   Call VerificarEventos(sDate)
   Call VerificarAtendimentos(sDate)
   Call VerificarVendas(sDate)
   Call VerificarTarefas(sDate)
   
   Call ResizeGrid
   
   mvarMe.ProgressBar1.Value = 0
   mvarMe.ProgressBar1.Scrolling = xtpProgressBarSmooth
   mvarMe.ProgressBar1.Enabled = False
   mvarMe.ProgressBar1.Visible = False
   
   
   With mvarMe.TaskDialog
      .Reset
      .MessageBoxStyle = True
      .WindowTitle = "Fechamento do Dia"
      .VerifyCheckState = True
      
      .FooterIcon = xtpTaskIconInformation
      .FooterText = "Para enviar o fechamento ao sócios é necessário estar conectado à Internet. Caso a conexão não esteja ativa, desmarque a opção acima e continue o fechamento normalmente."
      If nTotalErros = 0 Then
         'From General Tab
         '.EnableCallbackTimer = True
         '.ProgressBarStyle = xtpTaskProgressBarNormal
         '.AllowCancellation = True
         '.SetProgressBarRange 0, 5000
         
         
         .VerificationText = "Comunicação aos sócios"
         'xxx If nTotalAvisos = 0 Then
         If False Then
            .MainInstructionText = "Não há erros para Fechamento."
            .ContentText = "A seguir será realizado o fechamento das atividades e a cópia de segurança. Esta operação poderá levar alguns minutos, por favor aguarde."
         
            If mvarSys.MDI.Icon = 0 Then
               .MainIcon = xtpTaskIconShield
            Else
               .MainIcon = xtpTaskIconCustom
               .MainIconHandle = mvarSys.MDI.Icon
            End If
         Else
            .CommonButtons = 0
            .CommonButtons = xtpTaskButtonYes Or xtpTaskButtonNo
            
            .MainInstructionText = "Existem alguns comentários no fechamento." & vbNewLine & vbNewLine & "Deseja continuar?"
            .ContentText = "Os comentários existentes não impedem o fechamento do dia. Caso continue, o fechamento das atividades e a cópia de segurança será realizada normalmente."
            '.FooterText = "Os comentários existentes não impedem o fechamento do dia. Caso continue, o fechamento das atividades e a cópia de segurança será realizada normalmente. Lembrando que esta operação poderá levar alguns minutos, por favor aguarde."
            .ExpandedInformationText = "Total de Erros: " & StrZero(nTotalErros, 2) & vbNewLine & "Total de Avisos: " & StrZero(nTotalAvisos, 2)
            .ExpandedControlText = "Detalhes..."
            .MainIcon = xtpTaskIconWarning
         End If
                           
         .ShowDialog
               
      Else
         'From General Tab
         .MainInstructionText = "Foi encontrado alguns erros e o fechamento não poderá continuar."
         '.ContentText = ""
         .FooterText = "Por favor retorne a tela anterior, ajuste as sinalizações e execute a verificação novamente."
         .VerificationText = ""
         .VerifyCheckState = True
         
         'From Expanded Tab
         .ExpandedInformationText = "Total de Erros: " & StrZero(nTotalErros, 2) & vbNewLine & "Total de Avisos: " & StrZero(nTotalAvisos, 2)
         .ExpandedControlText = "Detalhes..."
         .CollapsedControlText = .ExpandedControlText
         .FooterIcon = xtpTaskIconInformation
         .MainIcon = xtpTaskIconError
                     
         .ShowDialog
      End If
      
   End With
         
   Screen.MousePointer = vbDefault
End Sub
Private Sub ResizeGrid()
   With mvarMe.GrdFecharDia
      .AllowColumnResize = True
      .Columns(0).Width = 4
      .Columns(1).Width = 2
      .Columns(2).Width = 30
      .AllowColumnResize = False
   End With
End Sub
Private Sub ExibirItens()
   Dim xRec As ReportRecord
   Dim xSubRec As ReportRecord
   Dim nVisible As Integer
   
   For Each xRec In mvarMe.GrdFecharDia.Records
      nVisible = 0
      For Each xSubRec In xRec.Childs
         If xSubRec.Item(1).Icon = Bandeira.Ok Then xSubRec.Visible = (mvarMe.ChkOk.Value = xtpChecked)
         If xSubRec.Item(1).Icon = Bandeira.Erro Then xSubRec.Visible = (mvarMe.ChkErros.Value = xtpChecked)
         If xSubRec.Item(1).Icon = Bandeira.Aviso Then xSubRec.Visible = (mvarMe.ChkAvisos.Value = xtpChecked)
         
         If Not xSubRec.Visible Then nVisible = nVisible + 1
      Next
      xRec.Visible = True
      If xRec.Childs.Count = nVisible Then
         If xRec.Item(0).Icon = Bandeira.Verde Then xRec.Visible = (mvarMe.ChkOk.Value = xtpChecked)
         If xRec.Item(0).Icon = Bandeira.Vermelha Then xRec.Visible = (mvarMe.ChkErros.Value = xtpChecked)
         If xRec.Item(0).Icon = Bandeira.Amarela Then xRec.Visible = (mvarMe.ChkAvisos.Value = xtpChecked)
      End If
   Next
   mvarMe.GrdFecharDia.Populate
End Sub
Private Sub LimparGrid()
   Dim xCol As ReportColumn
   Dim i    As Integer
   
   With mvarMe.GrdFecharDia
      .Icons = mvarMe.ImageManager1.Icons
      .ShowGroupBox = False
      .BorderStyle = xtpBorderFrame
      .AllowColumnRemove = False
      .AllowColumnReorder = False
      .AllowColumnResize = False
      .AllowColumnSort = False
      .AllowEdit = False
      .MultipleSelection = False
      .ShowHeaderRows = False

      
      .Records.DeleteAll
      .Columns.DeleteAll
      Set xCol = .Columns.Add(0, " ", 30, False)
      Set xCol = .Columns.Add(1, " ", 30, False)
      Set xCol = .Columns.Add(2, "Itens", 1, True)
      
      For i = 0 To .Columns.Count - 1
         .Columns(i).TreeColumn = (i = 0)
         .Columns(i).AutoSortWhenGrouped = True
         .Columns(i).AllowDrag = False
         .Columns(i).AllowRemove = False
         .Columns(i).AutoSize = True
         .Columns(i).Editable = False
      Next
      
      .GroupsOrder.Add .Columns(0)
      .Populate
   End With
End Sub
Private Sub VerificarEventos(pData As String)
   Dim xSubRec    As ReportRecord
   Dim xRecord    As ReportRecord
   Dim Sql        As String
   Dim MyRs       As Object
   Dim sIDAntes   As String
   Dim sGrupo     As String
   Dim i          As Integer
   Dim sCaption   As String
   Dim QtdIni     As Integer
   Dim QtdTotal   As Integer
   Dim QtdErro    As Integer
   Dim QtdAviso   As Integer
   Dim bNewId     As Boolean
   Dim nTreeColumn As Integer
   
   Sql = "Select E.IDLOJA, E.IDEVENTO, E.FLGCONFIRMADO, E.FLGCANCELADO, E.STARTDATETIME"
   Sql = Sql & ", C.IDCLIENTE, C.NOME, A.IDATENDIMENTO"
   Sql = Sql & " From OEVENTOAGENDA E"
   Sql = Sql & " Left Join OCLIENTE      C On E.IDLOJA=C.IDLOJA And E.IDCLIENTE=C.IDCLIENTE"
   Sql = Sql & " Left Join OATENDIMENTO  A On A.IDLOJA=E.IDLOJA And A.IDEVENTO=E.IDEVENTO"
   Sql = Sql & " Where CONVERT(VARCHAR(10), E.STARTDATETIME, 103 )=" & SqlDate(pData, 1)
   'Sql = Sql & " And E.ISMEETING = 1"
   'Sql = Sql & " And E.FLGCANCELADO <> 1"
   Sql = Sql & " Order by E.STARTDATETIME"
      
   sGrupo = "OEVENTOAGENDA"
   QtdIni = mvarMe.GrdFecharDia.Records.Count
   nTreeColumn = 0
   If mvarSys.xdb.AbreTabela(Sql, MyRs) Then
      sIDAntes = ""
      QtdTotal = 0
      QtdErro = 0
      QtdAviso = 0
      With MyRs
         While Not .EOF
            If sIDAntes <> .Fields("IDEVENTO") & "" Then
               QtdTotal = QtdTotal + 1
            
               sCaption = "Evento " & StrZero(.Fields("IDEVENTO") & "", 6)
               sCaption = sCaption & " - " & Format(.Fields("STARTDATETIME") & "", "HH:MM")
               sCaption = sCaption & " - " & Trim(Mid(.Fields("NOME") & "", 1, 35))
               
               Set xRecord = AddReportRecord(mvarMe.GrdFecharDia, Nothing, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
               xRecord.Expanded = True
               xRecord.Item(0).Icon = Bandeira.Verde
               If .Fields("FLGCANCELADO") = 1 Then
                  sCaption = "Evento " & StrZero(.Fields("IDEVENTO") & "", 6)
                  sCaption = sCaption & " - " & Format(.Fields("STARTDATETIME") & "", "HH:MM")
                  sCaption = sCaption & " - " & Trim(Mid(.Fields("NOME") & "", 1, 25))
                  sCaption = sCaption & "... (Cancelado)"
                  
                  xRecord.Item(2).Caption = sCaption
                  
               End If
               xRecord.Item(2).Hyperlinks.AddHyperlink Len("Evento "), 6
               Call SetTag(xRecord.Item(2), "TABELA", sGrupo)
            
               '**************
               '* ERROS
               '**************
               '*********
               '* Verificar Atendimentos
               If .Fields("IDATENDIMENTO") & "" = "" And .Fields("FLGCANCELADO") <> 1 Then
                  sCaption = "Sem atendimento associado."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               '*********
               '* Verificar Integridades
               If Val(.Fields("IDEVENTO") & "") < 0 Then
                  sCaption = "Sem integridade no campo [Identificador] (" & .Fields("IDEVENTO") & "" & ")"
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               
            End If
            
            sIDAntes = .Fields("IDEVENTO") & ""
            .MoveNext
            
            bNewId = .EOF
            If Not bNewId Then bNewId = (sIDAntes <> .Fields("IDATENDIMENTO") & "")
            If bNewId Then
               If xRecord.Item(0).Icon <> Bandeira.Vermelha Then
                  For Each xSubRec In xRecord.Childs
                     If xSubRec.Item(1).Icon = Bandeira.Aviso Then
                        If xRecord.Item(0).Icon <> Bandeira.Vermelha Then
                           xRecord.Item(0).Icon = Bandeira.Amarela
                        End If
                        QtdAviso = QtdAviso + 1
                     End If
                     If xSubRec.Item(1).Icon = Bandeira.Erro Or xSubRec.Item(1).Icon = Bandeira.Aviso Then
                        xRecord.Item(0).Icon = Bandeira.Vermelha
                        QtdErro = QtdErro + 1
                     End If
                  Next
               End If
               If xRecord.Item(0).Icon = Bandeira.Verde Then
                  xRecord.Childs.DeleteAll
                  xRecord.Item(2).Hyperlinks.RemoveAll
               End If
            End If
         Wend
      End With
   Else
      Set xRecord = AddReportRecord(mvarMe.GrdFecharDia, Nothing, Array("", "", "Sem eventos para analisar"), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
      xRecord.Expanded = False
      xRecord.Item(0).Icon = Bandeira.Verde
   End If
   
   '* Determina Grupo
   With mvarMe.GrdFecharDia
      For i = QtdIni To .Records.Count - 1
         Set xSubRec = .Records(i)
         With xSubRec.Item(0)
            If .GroupCaption = sGrupo Then
               .GroupCaption = ""
               .GroupCaption = .GroupCaption & StrZero(QtdTotal, IIf(QtdTotal = 0, 0, 2)) & " Evento" & IIf(QtdTotal = 1, "", "s")
               .GroupCaption = .GroupCaption & " / " & StrZero(QtdErro, IIf(QtdErro = 0, 0, 2)) & " Erro" & IIf(QtdErro = 1, "", "s")
               .GroupCaption = .GroupCaption & " / " & StrZero(QtdAviso, IIf(QtdAviso = 0, 0, 2)) & " Aviso" & IIf(QtdAviso = 1, "", "s")

               .GroupPriority = 1
            End If
         End With
      Next
      nTotalErros = nTotalErros + QtdErro
      nTotalAvisos = nTotalAvisos + QtdAviso
      .Populate
   End With
End Sub
Private Sub VerificarAtendimentos(pData As String)
   Dim xSubRec    As ReportRecord
   Dim xRecord    As ReportRecord
   Dim Sql        As String
   Dim MyRs       As Object
   Dim sIDAntes   As String
   Dim sGrupo     As String
   Dim i          As Integer
   Dim sCaption   As String
   Dim QtdIni     As Integer
   Dim QtdTotal   As Integer
   Dim QtdErro    As Integer
   Dim QtdAviso   As Integer
   Dim bNewId     As Boolean
   
   Dim bSessaoOK   As Boolean
   Dim bDirecaoOK  As Boolean
   Dim bFotoOK     As Boolean
   Dim bPotenciaOK As Boolean
   Dim bDisparosOK As Boolean
   Dim bTempoOK    As Boolean
   
   Dim bVendaOk    As Boolean
   Dim bExisteVenda As Boolean
         
   Sql = ""
   Sql = Sql & "Select A.IDLOJA, A.IDATENDIMENTO, A.IDEVENTO, A.IDCLIENTE, A.IDFUNCIONARIO"
   Sql = Sql & ", A.IDMAQUINA, A.IDSALA, A.HHINI, A.HHFIM, S.IDSESSAO"
   Sql = Sql & ", S.IDTPSERVICO, S.IDTPTRATAMENTO, S.IDAREA, S.IDDIRECAO, S.FOTOTIPO"
   Sql = Sql & ", S.POTENCIA, S.DISPAROS, S.NUMSESSAO, S.TEMPOSESSAO, V.IDVENDA"
   Sql = Sql & ", P.IDPROD, C.IDCLIENTE, C.NOME, E.IDEVENTO, P2.EVENDA"
   Sql = Sql & " From OATENDIMENTO               A"
   Sql = Sql & " Join OCLIENTE                   C On A.IDLOJA =C.IDLOJA  And A.IDCLIENTE=C.IDCLIENTE"
   Sql = Sql & " Left Join OSESSAO               S On A.IDLOJA =S.IDLOJA  And A.IDATENDIMENTO=S.IDATENDIMENTO"
   Sql = Sql & " Left Join OATENDIMENTO_VENDA    V On A.IDLOJA =V.IDLOJA  And A.IDATENDIMENTO=V.IDATENDIMENTO"
   Sql = Sql & " Left Join OATENDIMENTO_PRODUTO AP On A.IDLOJA =AP.IDLOJA And A.IDATENDIMENTO=AP.IDATENDIMENTO"
   Sql = Sql & " Left Join SPRODUTO              P On AP.IDPROD=P.IDPROD  And P.EVENDA=1"
   Sql = Sql & " Left Join OEVENTOAGENDA         E On E.IDLOJA =A.IDLOJA  And E.IDEVENTO=A.IDEVENTO"
   Sql = Sql & " Left Join OTRATAMENTO_PROD      TP On TP.IDTPSERVICO=S.IDTPSERVICO"
   Sql = Sql & " Left Join SPRODUTO             P2 On TP.IDPROD=P2.IDPROD And P2.EVENDA=1"
   Sql = Sql & " Where CONVERT(VARCHAR(10), A.DTATEND, 103 )=" & SqlDate(pData, 1)
   Sql = Sql & " And A.SITATEND Not In ('9X')" ', '9C')"
   Sql = Sql & " Order by A.DTATEND, A.HHINI"
      
   sGrupo = "OATENDIMENTO"
   QtdIni = mvarMe.GrdFecharDia.Records.Count
   If mvarSys.xdb.AbreTabela(Sql, MyRs) Then
      sIDAntes = ""
      QtdTotal = 0
      QtdErro = 0
      QtdAviso = 0
      
      bSessaoOK = True
      bDirecaoOK = True
      bFotoOK = True
      bPotenciaOK = True
      bDisparosOK = True
      bTempoOK = True
      
      bVendaOk = True
      bExisteVenda = False
      
      Dim nTreeColumn As Integer
      nTreeColumn = 0
      With MyRs
         While Not .EOF
         
            If sIDAntes <> .Fields("IDATENDIMENTO") & "" Then
               QtdTotal = QtdTotal + 1
            
               sCaption = "Atendimento " & StrZero(.Fields("IDATENDIMENTO") & "", 6)
               sCaption = sCaption & " - " & Format(.Fields("HHINI") & "", "HH:MM")
               sCaption = sCaption & " - " & Trim(Mid(.Fields("NOME") & "", 1, 35))
               
               Set xRecord = AddReportRecord(mvarMe.GrdFecharDia, Nothing, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
               xRecord.Expanded = True
               xRecord.Item(0).Icon = Bandeira.Verde
               xRecord.Item(2).Hyperlinks.AddHyperlink Len("Atendimento "), 6
               Call SetTag(xRecord.Item(2), "TABELA", sGrupo)
'mvarMe.GrdFecharDia.Populate
               '**************
               '* ERROS
               '**************
               '* Verificar Cliente
               If xVal(.Fields("IDCLIENTE") & "") <= 0 Then
                  sCaption = "Sem cliente definido."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               
               '* Verificar Sala
               If xVal(.Fields("IDSALA") & "") <= 0 Then
                  sCaption = "Sem sala definida."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               
               '* Verificar Máquina
               If xVal(.Fields("IDMAQUINA") & "") <= 0 Then
                  sCaption = "Sem máquina definida."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               
               '* Verificar Operadora
               If xVal(.Fields("IDFUNCIONARIO") & "") <= 0 Then
                  sCaption = "Sem funcionária definida."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               
               '* Verificar Evento
               If xVal(.Fields("IDEVENTO") & "") <= 0 Then
                  sCaption = "Sem evento associado."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               
               '* Verificar Venda
               bVendaOk = (xVal(.Fields("IDVENDA") & "") > 0)
               bExisteVenda = (xVal(.Fields("EVENDA") & "") = 1)
               
               
               '* Verificar Registro de Tempo
               If Not IsDate(.Fields("HHINI") & "") Or Not IsDate(.Fields("HHINI") & "") Then
                  sCaption = "Inconsistência no horário."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
            End If
               
            If sIDAntes = .Fields("IDATENDIMENTO") & "" Or Trim(sIDAntes) = "" Then
               bExisteVenda = (xVal(.Fields("EVENDA") & "") = 1)
               
               '**************
               '* AVISOS
               '**************
               '* Verificar Nº da Sessao
               bSessaoOK = IIf(bSessaoOK, (xVal(.Fields("IDSESSAO") & "") > 0), bSessaoOK)
               '* Verificar Nº Direção
               bDirecaoOK = IIf(bSessaoOK, (xVal(.Fields("IDSESSAO") & "") > 0), bSessaoOK)
               '* Verificar Nº Fototipo
               bFotoOK = IIf(bSessaoOK, (xVal(.Fields("IDSESSAO") & "") > 0), bSessaoOK)
               '* Verificar Nº Potência
               bPotenciaOK = IIf(bSessaoOK, (xVal(.Fields("IDSESSAO") & "") > 0), bSessaoOK)
               '* Verificar Nº Disparos
               bDisparosOK = IIf(bSessaoOK, (xVal(.Fields("IDSESSAO") & "") > 0), bSessaoOK)
               '* Verificar Nº Tempo
               bTempoOK = IIf(bSessaoOK, (xVal(.Fields("IDSESSAO") & "") > 0), bSessaoOK)
            End If
            
            sIDAntes = .Fields("IDATENDIMENTO") & ""
            .MoveNext
            
            bNewId = .EOF
            If Not bNewId Then bNewId = (sIDAntes <> .Fields("IDATENDIMENTO") & "")
            If bNewId Then
               If bExisteVenda And Not bVendaOk Then
                  sCaption = "Sem venda associada."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
            
               If Not bSessaoOK Then
                  sCaption = "Erro no campo Sessão."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               If Not bDirecaoOK Then
                  sCaption = "Sem integridade no campo [Direção]."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Aviso
               End If
               If Not bFotoOK Then
                  sCaption = "Sem integridade no campo [Fototipo]."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Aviso
               End If
               If Not bPotenciaOK Then
                  sCaption = "Sem integridade no campo [Potência]."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Aviso
               End If
               If Not bDisparosOK Then
                  sCaption = "Sem integridade no campo [Nº Disparos]."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Aviso
               End If
               If Not bTempoOK Then
                  sCaption = "Sem integridade no campo [Tempo da Sessão]."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Aviso
               End If
                              
               bSessaoOK = True
               bDirecaoOK = True
               bFotoOK = True
               bPotenciaOK = True
               bDisparosOK = True
               bTempoOK = True
               
               bVendaOk = True
               bExisteVenda = False
                              
               If xRecord.Item(0).Icon <> Bandeira.Vermelha Then
                  For Each xSubRec In xRecord.Childs
                     If xSubRec.Item(1).Icon = Bandeira.Aviso Then
                        If xRecord.Item(0).Icon <> Bandeira.Vermelha Then
                           xRecord.Item(0).Icon = Bandeira.Amarela
                        End If
                        QtdAviso = QtdAviso + 1
                     End If
                     If xSubRec.Item(1).Icon = Bandeira.Erro Then
                        xRecord.Item(0).Icon = Bandeira.Vermelha
                        QtdErro = QtdErro + 1
                     End If
                  Next
               End If
               If xRecord.Item(0).Icon = Bandeira.Verde Then
                  xRecord.Childs.DeleteAll
                  xRecord.Item(2).Hyperlinks.RemoveAll
               End If
            End If
         Wend
      End With
   Else
      sCaption = "Sem atendimentos para analisar"
      Set xRecord = AddReportRecord(mvarMe.GrdFecharDia, Nothing, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
      xRecord.Expanded = False
      xRecord.Item(0).Icon = Bandeira.Verde
   End If
   
   '* Determina Grupo
   With mvarMe.GrdFecharDia
      For i = QtdIni To .Records.Count - 1
         Set xSubRec = .Records(i)
         With xSubRec.Item(0)
            If .GroupCaption = sGrupo Then
               .GroupCaption = ""
               .GroupCaption = .GroupCaption & StrZero(QtdTotal, IIf(QtdTotal = 0, 0, 2)) & " Atendimento" & IIf(QtdTotal = 1, "", "s")
               .GroupCaption = .GroupCaption & " / " & StrZero(QtdErro, IIf(QtdErro = 0, 0, 2)) & " Erro" & IIf(QtdErro = 1, "", "s")
               .GroupCaption = .GroupCaption & " / " & StrZero(QtdAviso, IIf(QtdAviso = 0, 0, 2)) & " Aviso" & IIf(QtdAviso = 1, "", "s")

               .GroupPriority = 1
            End If
         End With
      Next
      nTotalErros = nTotalErros + QtdErro
      nTotalAvisos = nTotalAvisos + QtdAviso
      .Populate
   End With
End Sub
Private Sub VerificarTarefas(pData As String)
   With mvarMe.GrdFecharDia
   End With
End Sub
Private Sub VerificarVendas(pData As String)
   Dim xSubRec    As ReportRecord
   Dim xRecord    As ReportRecord
   Dim Sql        As String
   Dim MyRs       As Object
   Dim RsAux       As Object
   Dim sIDAntes   As String
   Dim sGrupo     As String
   Dim i          As Integer
   Dim sCaption   As String
   Dim QtdIni     As Integer
   Dim QtdTotal   As Integer
   Dim QtdErro    As Integer
   Dim QtdAviso   As Integer
   Dim bNewId     As Boolean
   
   Dim bQtdVendaOK   As Boolean
   Dim bPgtosOK      As Boolean
   Dim bFlanOK       As Boolean
   Dim bFlanExiste   As Boolean
   Dim nQdProd       As Double
   Dim nQdserv       As Double
   Dim nValor        As Double
         
   Sql = ""
   Sql = Sql & "Select V.IDLOJA, V.IDVENDA, V.IDCLIENTE, V.IDFUNCIONARIO, V.SITVENDA"
   Sql = Sql & ", C.NOME, V.VLVENDA, V.VLPGTO, V.VLDESC"
   Sql = Sql & " From CVENDA     V"
   'Sql = Sql & " Join CITENVENDA I On V.IDLOJA =I.IDLOJA  And V.IDVENDA=I.IDVENDA"
   Sql = Sql & " Join OCLIENTE   C On V.IDLOJA =C.IDLOJA  And V.IDCLIENTE=C.IDCLIENTE"
   'Sql = Sql & " Join SPRODUTO   P On I.IDPROD =P.IDPROD  And P.EVENDA=1"
   'Sql = Sql & " Left Join OATENDIMENTO_VENDA   AV On V.IDLOJA =AV.IDLOJA  And V.IDVENDA=AV.IDVENDA"
   'Sql = Sql & " Left Join OATENDIMENTO          A On AV.IDLOJA =A.IDLOJA  And AV.IDATENDIMENTO=A.IDATENDIMENTO"
   'Sql = Sql & " Left Join OSESSAO               S On A.IDLOJA =S.IDLOJA   And A.IDATENDIMENTO=S.IDATENDIMENTO"
   'Sql = Sql & " Left Join OATENDIMENTO_PRODUTO AP On A.IDLOJA =AP.IDLOJA  And A.IDATENDIMENTO=AP.IDATENDIMENTO"
   'Sql = Sql & " Left Join SPRODUTO              P On AP.IDPROD=P.IDPROD   And P.EVENDA=1"
   'Sql = Sql & " Left Join OTRATAMENTO_PROD     TP On TP.IDTPSERVICO=S.IDTPSERVICO"
   'Sql = Sql & " Left Join SPRODUTO             P2 On TP.IDPROD=P2.IDPROD  And P2.EVENDA=1"
   'Sql = Sql & " Left Join CPGTOSVENDA          PG On V.IDLOJA =PG.IDLOJA  And V.IDVENDA=PG.IDVENDA"
   Sql = Sql & " Where CONVERT(VARCHAR(10), V.DTVENDA, 103 )=" & SqlDate(pData, 1)
   Sql = Sql & " And V.SITVENDA Not In ('9X')"
   Sql = Sql & " Order by V.DTVENDA, V.IDVENDA"
      
   sGrupo = "CVENDA"
   QtdIni = mvarMe.GrdFecharDia.Records.Count
   If mvarSys.xdb.AbreTabela(Sql, MyRs) Then
      sIDAntes = ""
      QtdTotal = 0
      QtdErro = 0
      QtdAviso = 0
      
      bQtdVendaOK = True
      
      Dim nTreeColumn As Integer
      nTreeColumn = 0
      With MyRs
         While Not .EOF
         
            If sIDAntes <> .Fields("IDVENDA") & "" Then
               QtdTotal = QtdTotal + 1
            
               sCaption = "Venda " & StrZero(.Fields("IDVENDA") & "", 6)
               sCaption = sCaption & " - R$" & ValBr(.Fields("VLVENDA") & "")
               sCaption = sCaption & " - " & Trim(Mid(.Fields("NOME") & "", 1, 35))
               
               Set xRecord = AddReportRecord(mvarMe.GrdFecharDia, Nothing, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
               xRecord.Expanded = True
               xRecord.Item(0).Icon = Bandeira.Verde
               xRecord.Item(2).Hyperlinks.AddHyperlink Len("Venda "), 6
               Call SetTag(xRecord.Item(2), "TABELA", sGrupo)
               '**************
               '* ERROS
               '**************
               '* Verificar Cliente
               If xVal(.Fields("IDCLIENTE") & "") <= 0 Then
                  sCaption = "Sem cliente definido."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               
               '* Verificar Operadora
'               If xVal(.Fields("IDFUNCIONARIO") & "") <= 0 Then
'                  sCaption = "Sem funcionária definida."
'                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
'                  xSubRec.Item(1).Icon = Bandeira.Erro
'               End If
                                                            
               '* Verificar Valor de Pagamento
               If xVal(.Fields("VLPGTO") & "") <> xVal(.Fields("VLVENDA") & "") - xVal(.Fields("VLDESC") & "") Then
                  sCaption = "Valor de pagamento (R$" & ValBr(xVal(.Fields("VLPGTO") & "")) & ") difere do à pagar (R$" & ValBr(xVal(.Fields("VLVENDA") & "") - xVal(.Fields("VLDESC") & "")) & ")."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
                                                            
               '* Verificar Pagamentos
               Sql = "Select Sum(VLPGTO) As [VLPGTO]"
               Sql = Sql & " From CPGTOSVENDA"
               Sql = Sql & " Where IDLOJA=" & SqlNum(.Fields("IDLOJA") & "")
               Sql = Sql & " And IDVENDA=" & SqlNum(.Fields("IDVENDA") & "")
               If mvarSys.xdb.AbreTabela(Sql, RsAux) Then
                  bPgtosOK = (xVal(.Fields("VLPGTO") & "") = xVal(RsAux("VLPGTO") & ""))
               Else
                  bPgtosOK = (xVal(.Fields("VLPGTO") & "") = 0)
               End If
               If Not bPgtosOK Then
                  sCaption = "Valor pago (R$" & ValBr(xVal(.Fields("VLPGTO") & "")) & ") difere do à pagar (R$" & ValBr(xVal(RsAux("VLPGTO"))) & ")."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If

                              
               '* Verificar Atendimento
               Sql = "Select A.IDATENDIMENTO"
               Sql = Sql & " From OATENDIMENTO     A"
               Sql = Sql & " Join OATENDIMENTO_VENDA   AV On A.IDLOJA =AV.IDLOJA  And A.IDATENDIMENTO=AV.IDATENDIMENTO"
               Sql = Sql & " Where AV.IDLOJA=" & SqlNum(.Fields("IDLOJA") & "")
               Sql = Sql & " And AV.IDVENDA=" & SqlNum(.Fields("IDVENDA") & "")
               If Not mvarSys.xdb.AbreTabela(Sql, RsAux) Then
                  '* Verificar se Existe Atendimento
                  sCaption = "Venda sem atendimento associado."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Aviso
               Else
                  '* Verificar Quantidade de Venda

                  Sql = "Select sum((IsNull(AP.QTDCOMPRA,0)*IsNull(P.VLVENDA,0))+(IsNull(P2.IDPROD,0)*IsNull(P2.VLVENDA,0))) As [Valor]"
                  Sql = Sql & ", IsNull(Count(P.IDPROD),0) As [PROD], IsNull(Count( P2.IDPROD),0) As [SERV]"
                  Sql = Sql & " From OATENDIMENTO     A"
                  Sql = Sql & " Join OATENDIMENTO_VENDA   AV On A.IDLOJA =AV.IDLOJA  And A.IDATENDIMENTO=AV.IDATENDIMENTO"
                  Sql = Sql & " Left Join OSESSAO               S On A.IDLOJA =S.IDLOJA   And A.IDATENDIMENTO=S.IDATENDIMENTO"
                  Sql = Sql & " Left Join OATENDIMENTO_PRODUTO AP On A.IDLOJA =AP.IDLOJA  And A.IDATENDIMENTO=AP.IDATENDIMENTO"
                  Sql = Sql & " Left Join SPRODUTO              P On AP.IDPROD=P.IDPROD   And P.EVENDA=1"
                  Sql = Sql & " Left Join OTRATAMENTO_PROD     TP On TP.IDTPSERVICO=S.IDTPSERVICO"
                  Sql = Sql & " Left Join SPRODUTO             P2 On TP.IDPROD=P2.IDPROD  And P2.EVENDA=1"
                  Sql = Sql & " Where AV.IDLOJA=" & SqlNum(.Fields("IDLOJA") & "")
                  Sql = Sql & " And AV.IDVENDA=" & SqlNum(.Fields("IDVENDA") & "")
                  Sql = Sql & " Group By A.IDATENDIMENTO"
                  If mvarSys.xdb.AbreTabela(Sql, RsAux) Then
                     nQdProd = RsAux("PROD")
                     nQdserv = RsAux("SERV")
                     nValor = RsAux("Valor")
                     
                     Sql = "Select IsNull(Sum(QTDVENDA),0) as [QTDVENDA],  IsNull(Sum(QTDVENDA*VLUNIT),0) as [VALOR]"
                     Sql = Sql & " From CITENSVENDA "
                     Sql = Sql & " Where IDLOJA =" & SqlNum(.Fields("IDLOJA") & "")
                     Sql = Sql & " And   IDVENDA=" & SqlNum(.Fields("IDVENDA") & "")
                     bQtdVendaOK = False
                     If mvarSys.xdb.AbreTabela(Sql, RsAux) Then
                        If nQdProd + nQdserv = Val(RsAux("QTDVENDA")) Or nValor = Val(RsAux("VALOR")) Then
                           bQtdVendaOK = True
                        End If
                     End If
                  End If
                  If Not bQtdVendaOK Then
                     '* Verificar Quantidade de Venda
                     sCaption = "Quantidade/Valor de venda não cenfere com a quantidade no atendimento."
                     Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                     xSubRec.Item(1).Icon = Bandeira.Erro
                  End If
               End If
                              
               '* Verificar Situação da venda
               If .Fields("SITVENDA") & "" <> "10" Then
                  sCaption = "Venda sem fechamento."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
                              
               '* Verificar Lançamento Financeiro
               Sql = "Select P.IDVENDA, P.IDPGTO, P.VLPGTO, F.IDVENDA, F.VALOR"
               Sql = Sql & " From CPGTOSVENDA P"
               Sql = Sql & " Left Join FLAN F On P.IDLOJA =F.IDLOJA And P.IDVENDA=F.IDVENDA And P.IDPGTO=F.IDPGTO"
               Sql = Sql & " Where P.IDLOJA=" & SqlNum(.Fields("IDLOJA") & "")
               Sql = Sql & " And P.IDVENDA=" & SqlNum(.Fields("IDVENDA") & "")
               bFlanOK = False
               bFlanExiste = False
               
               bFlanExiste = mvarSys.xdb.AbreTabela(Sql, RsAux)
               If bFlanExiste Then
                  nValor = 0
                  While Not RsAux.EOF
                     nValor = nValor + xVal(RsAux("Valor") & "")
                     bFlanOK = (xVal(RsAux("VLPGTO")) = nValor)
                     RsAux.MoveNext
                  Wend
                  If Not bFlanOK Then
                     sCaption = "Valor do pagamento (R$" & ValBr(xVal(.Fields("VLPGTO") & "")) & ") difere do lançamento (R$" & ValBr(nValor) & ")."
                     Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                     xSubRec.Item(1).Icon = Bandeira.Erro
                  End If
                  
               Else
                  sCaption = "Pagamentos sem lançamento Financeiro."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
            End If
               
'            If sIDAntes = .Fields("IDATENDIMENTO") & "" Or Trim(sIDAntes) = "" Then
'               bExisteVenda = (xVal(.Fields("EVENDA") & "") = 1)
'
'               '**************
'               '* AVISOS
'               '**************
'               '* Verificar Nº da Sessao
'               bSessaoOK = IIf(bSessaoOK, (xVal(.Fields("IDSESSAO") & "") > 0), bSessaoOK)
'               '* Verificar Nº Direção
'               bDirecaoOK = IIf(bSessaoOK, (xVal(.Fields("IDSESSAO") & "") > 0), bSessaoOK)
'               '* Verificar Nº Fototipo
'               bFotoOK = IIf(bSessaoOK, (xVal(.Fields("IDSESSAO") & "") > 0), bSessaoOK)
'               '* Verificar Nº Potência
'               bPotenciaOK = IIf(bSessaoOK, (xVal(.Fields("IDSESSAO") & "") > 0), bSessaoOK)
'               '* Verificar Nº Disparos
'               bDisparosOK = IIf(bSessaoOK, (xVal(.Fields("IDSESSAO") & "") > 0), bSessaoOK)
'               '* Verificar Nº Tempo
'               bTempoOK = IIf(bSessaoOK, (xVal(.Fields("IDSESSAO") & "") > 0), bSessaoOK)
'            End If
            
            sIDAntes = .Fields("IDVENDA") & ""
            .MoveNext
            
            bNewId = .EOF
            If Not bNewId Then bNewId = (sIDAntes <> .Fields("IDVENDA") & "")
            If bNewId Then
'               If bExisteVenda And Not bVendaOk Then
'                  sCaption = "Sem venda associada."
'                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
'                  xSubRec.Item(1).Icon = Bandeira.Erro
'               End If
'
'               If Not bSessaoOK Then
'                  sCaption = "Erro no campo Sessão."
'                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
'                  xSubRec.Item(1).Icon = Bandeira.Erro
'               End If
'               If Not bDirecaoOK Then
'                  sCaption = "Sem integridade no campo [Direção]."
'                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
'                  xSubRec.Item(1).Icon = Bandeira.Aviso
'               End If
'               If Not bFotoOK Then
'                  sCaption = "Sem integridade no campo [Fototipo]."
'                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
'                  xSubRec.Item(1).Icon = Bandeira.Aviso
'               End If
'               If Not bPotenciaOK Then
'                  sCaption = "Sem integridade no campo [Potência]."
'                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
'                  xSubRec.Item(1).Icon = Bandeira.Aviso
'               End If
'               If Not bDisparosOK Then
'                  sCaption = "Sem integridade no campo [Nº Disparos]."
'                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
'                  xSubRec.Item(1).Icon = Bandeira.Aviso
'               End If
'               If Not bTempoOK Then
'                  sCaption = "Sem integridade no campo [Tempo da Sessão]."
'                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
'                  xSubRec.Item(1).Icon = Bandeira.Aviso
'               End If
                              
               If xRecord.Item(0).Icon <> Bandeira.Vermelha Then
                  For Each xSubRec In xRecord.Childs
                     If xSubRec.Item(1).Icon = Bandeira.Aviso Then
                        If xRecord.Item(0).Icon <> Bandeira.Vermelha Then
                           xRecord.Item(0).Icon = Bandeira.Amarela
                        End If
                        QtdAviso = QtdAviso + 1
                     End If
                     If xSubRec.Item(1).Icon = Bandeira.Erro Then
                        xRecord.Item(0).Icon = Bandeira.Vermelha
                        QtdErro = QtdErro + 1
                     End If
                  Next
               End If
               If xRecord.Item(0).Icon = Bandeira.Verde Then
                  xRecord.Childs.DeleteAll
                  xRecord.Item(2).Hyperlinks.RemoveAll
               End If
            End If
         Wend
      End With
   Else
      sCaption = "Sem vendas para analisar"
      Set xRecord = AddReportRecord(mvarMe.GrdFecharDia, Nothing, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
      xRecord.Expanded = False
      xRecord.Item(0).Icon = Bandeira.Verde
   End If
   
   '* Determina Grupo
   With mvarMe.GrdFecharDia
      For i = QtdIni To .Records.Count - 1
         Set xSubRec = .Records(i)
         With xSubRec.Item(0)
            If .GroupCaption = sGrupo Then
               .GroupCaption = ""
               .GroupCaption = .GroupCaption & StrZero(QtdTotal, IIf(QtdTotal = 0, 0, 2)) & " Venda" & IIf(QtdTotal = 1, "", "s")
               .GroupCaption = .GroupCaption & " / " & StrZero(QtdErro, IIf(QtdErro = 0, 0, 2)) & " Erro" & IIf(QtdErro = 1, "", "s")
               .GroupCaption = .GroupCaption & " / " & StrZero(QtdAviso, IIf(QtdAviso = 0, 0, 2)) & " Aviso" & IIf(QtdAviso = 1, "", "s")

               .GroupPriority = 1
            End If
         End With
      Next
      nTotalErros = nTotalErros + QtdErro
      nTotalAvisos = nTotalAvisos + QtdAviso
      .Populate
   End With
End Sub
Private Sub FecharVendas(pData As String)
   With mvarMe.GrdFecharDia
   End With
End Sub

Private Sub FecharAtendimentos(pData As String)
   With mvarMe.GrdFecharDia
   End With
End Sub
Private Sub InformarSocios(pData As String)
   With mvarMe.GrdFecharDia
   End With
End Sub
Private Sub GerarBackup()
   With mvarMe.GrdFecharDia
   End With
End Sub
Private Sub ArmazenarBackup()
   With mvarMe.GrdFecharDia
   End With
End Sub

Private Sub mvarMe_TaskDialogButtonClicked(ByVal Id As Long, CloseDialog As Variant)
   Select Case Id
      Case xtpTaskButtonOk, xtpTaskButtonYes '1,2
         If nTotalErros = 0 Then
            'mvarMe.TaskDialog.ProgressBarStyle = xtpTaskProgressBarMarquee
            'mvarMe.TaskDialog.StartProgressBarMarquee True, 10
'            bExeFechamento = True
'            With mvarMe.TaskFlood
'               .WindowTitle = "Error Sample (page 2)"
'               .ContentText = "Finalizando Dia"
'               .AllowCancellation = True
'               .SetProgressBarRange 0, 5000
'               .EnableCallbackTimer = True
'               .MessageBoxStyle = True
'               .ProgressBarStyle = xtpTaskProgressBarMarquee
'               .ShowDialog
'            End With
'            bExeFechamento
            CloseDialog = True
            Call FinalizarFechamento
         End If
   End Select
            
End Sub
Private Sub FinalizarFechamento()
   Dim sDate As String
   sDate = Format(mvarMe.CmbDTFECHAR.Value, "dd/mm/yyyy")

   'Call FecharVendas(sDate)
   Call FecharAtendimentos(sDate)
   Call GerarBackup
   
   If mvarMe.TaskDialog.VerifyCheckState Then
      Call InformarSocios(sDate)
      Call ArmazenarBackup
   End If

End Sub

Private Sub mvarMe_TaskFloodConstructed()
   mvarMe.TaskFlood.StartProgressBarMarquee True, 80
End Sub

'Private Sub mvarMe_TaskDialogTimer(ByVal MilliSeconds As Long, Reset As Variant)
'End Sub
Private Sub mvarMe_TaskFloodTimer(ByVal MilliSeconds As Long, Reset As Variant)
'    If 5000 >= MilliSeconds Then
'        Dim text As String
'        text = "Sending... " & Int(MilliSeconds / 5000 * 100) & "%"
'        taskDiaogFeedback.MainInstructionText = text
'
'        taskDiaogFeedback.ProgressBarPos = MilliSeconds
'        taskDiaogFeedback.SetProgressBarRange 0, 5000
'    Else
'        taskDiaogFeedback.MainInstructionText = "Thanks for the feedback!"
'        taskDiaogFeedback.ContentText = "Our developers will get right on that..."
'        taskDiaogFeedback.ProgressBarPos = 5000
'        taskDiaogFeedback.EnableCallbackTimer = False
'    End If
   If bExeFechamento Then
      Call FinalizarFechamento
      bExeFechamento = False
   End If
    
End Sub
