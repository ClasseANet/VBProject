VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NG_Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim gDebug As Boolean
Dim WithEvents mvarCalendar   As CalendarControl
Attribute mvarCalendar.VB_VarHelpID = -1
Dim WithEvents oFrmEvt        As FrmEventoCal
Attribute oFrmEvt.VB_VarHelpID = -1
Dim CollMes                   As Collection

Private mvarSys         As Object
Private mvarRsEvent     As Object
Private mvarIDCOLIGADA  As Integer
Private mvarIDLOJA      As Integer
Private mvarIDAGENDA    As Integer
'Private mvarIDCLIENTE   As Long
Private mvarTlEvt       As TL_Evento
Private mvarFLGAGENDADO As Integer
Private mvarFLGCANCELADO As Integer
Public Property Let FLGCANCELADO(ByVal vData As Integer)
   mvarFLGCANCELADO = vData
End Property
Public Property Get FLGCANCELADO() As Integer
   FLGCANCELADO = mvarFLGCANCELADO
End Property
Public Property Let FLGAGENDADO(ByVal vData As Integer)
   mvarFLGAGENDADO = vData
End Property
Public Property Get FLGAGENDADO() As Integer
   FLGAGENDADO = mvarFLGAGENDADO
End Property
Public Property Set TlEvt(ByVal vData As Object)
   Set mvarTlEvt = vData
   Set oFrmEvt = mvarTlEvt.Form
   'oFrmEvt.TxtIDCLIENTE.Text = 0
   'mvarIDCLIENTE = 0
End Property
Public Property Get TlEvt() As Object
   Set TlEvt = mvarTlEvt
End Property
Public Property Set CalendarControl(ByVal vData As Object)
   Set mvarCalendar = vData
End Property
'Public Property Get IDCLIENTE() As Long
'   IDCLIENTE = mvarIDCLIENTE
'End Property
'Public Property Let IDCLIENTE(ByVal vData As Long)
'   mvarIDCLIENTE = vData
'End Property
Public Property Let IDAGENDA(ByVal vData As Integer)
   mvarIDAGENDA = vData
End Property
Public Property Get IDAGENDA() As Integer
   IDAGENDA = mvarIDAGENDA
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
    mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDCOLIGADA(ByVal vData As Integer)
   mvarIDCOLIGADA = vData
End Property
Public Property Get IDCOLIGADA() As Integer
   IDCOLIGADA = mvarIDCOLIGADA
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   
   mvarIDCOLIGADA = xVal(mvarSys.Propriedades("IDCOLIGADA"))
   mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
   mvarIDAGENDA = mvarSys.Propriedades("IDAGENDA")
End Property
Public Property Get Sys() As Object
   Set Sys = mvarSys
End Property
Private Sub Class_Initialize()
   gDebug = False
   mvarIDCOLIGADA = 1
   mvarIDLOJA = 1
   mvarIDAGENDA = 1
   mvarFLGAGENDADO = 1
   mvarFLGCANCELADO = 1
End Sub

Private Sub mvarCalendar_BeforeEditOperation(ByVal OpParams As XtremeCalendarControl.CalendarEditOperationParameters, bCancelOperation As Boolean)
   If OpParams.Operation = xtpCalendarEO_DragCopy Or OpParams.Operation = xtpCalendarEO_DragMove Then
      If OpParams.DraggingEvent.AllDayEvent <> OpParams.DraggingEventNew.AllDayEvent Then
         bCancelOperation = True
      Else
         'bCancelOperation = Not VerificaEvento(OpParams.EventViews.ViewEvent(0).Event, False)
      End If
   ElseIf Not InArray(OpParams.Operation, Array(xtpCalendarEO_Copy, xtpCalendarEO_DragCopy)) Then
      If Not OpParams.EventViews Is Nothing Then
         If OpParams.EventViews.Count > 0 Then
            If InArray(OpParams.Operation, Array(xtpCalendarEO_Cut, xtpCalendarEO_DeleteEvent, xtpCalendarEO_DeleteSelectedEvents)) Then
               bCancelOperation = Not VerificaEvento(OpParams.EventViews.ViewEvent(0).Event, True)
            Else
               bCancelOperation = Not VerificaEvento(OpParams.EventViews.ViewEvent(0).Event, False)
            End If
         End If
      End If
   End If
End Sub
Private Sub mvarCalendar_DoCreateEvent(ByVal pEvent As XtremeCalendarControl.CalendarEvent, NewEventID As Long, bResult As Boolean)
   Dim oTbEvento  As Object 'TB_OEVENTOAGENDA
   Dim nMouse     As Integer
   
   On Error GoTo TrataErro
   
   bResult = False
   nMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   If Not mvarSys.xDb.Conectado Then GoTo Saida
   If Not VerificaEvento(pEvent, False) Then GoTo Saida

   Set oTbEvento = GetNewTbEvent
      
   Call MEventoParaClasse(pEvent, oTbEvento)

   If Not oTbEvento Is Nothing Then
      bResult = MSalvarEvento(oTbEvento)
      
      If bResult Then
         NewEventID = oTbEvento.IDEVENTO
         If MSalvarSessao(oTbEvento) Then
            Call mvarSys.xDb.Executa(QryRemarcado((oTbEvento.FLGCANCELADO = 0), oTbEvento.IDLOJA, oTbEvento.IDCLIENTE, oTbEvento.StartDateTime))
         End If
      Else
         Call ExibirInformacao("Erro ao gravar evento. [DoCreate]")
      End If
   End If

   GoTo Saida
TrataErro:
   Call ShowError
Saida:
    Screen.MousePointer = nMouse
End Sub
Private Sub mvarCalendar_DoCreateRPattern(ByVal pPattern As XtremeCalendarControl.CalendarRecurrencePattern, NewPatternID As Long, bResult As Boolean)
   Dim oTbEventoRec  As Object ' TB_OEVENTOREC
   Dim nMouse        As Integer

   On Error GoTo TrataErro
   
   nMouse = Screen.MousePointer
   bResult = False
   Screen.MousePointer = vbHourglass
   If Not mvarSys.xDb.Conectado Then GoTo Saida
      
   Set oTbEventoRec = GetNewTbEventR
   Call MEventoRecParaClasse(pPattern, oTbEventoRec)
   
   If Not oTbEventoRec Is Nothing Then
      bResult = oTbEventoRec.Salvar
      If bResult Then
         NewPatternID = oTbEventoRec.IDEVENTOREC
      Else
         Call ExibirInformacao("Erro ao gravar evento. [DoCreateR]")
      End If
   End If
   
   GoTo Saida
TrataErro:
   Call ShowError
Saida:
    Screen.MousePointer = nMouse
End Sub
Private Sub mvarCalendar_DoDeleteEvent(ByVal pEvent As XtremeCalendarControl.CalendarEvent, bResult As Boolean)
   Dim TbEvt   As Object 'TB_OEVENTOAGENDA
   Dim TbServ  As Object 'TB_OSERVICOEVT
   
   Dim oQueries As Collection
   Dim nMouse     As Integer
   nMouse = Screen.MousePointer
   
   On Error GoTo TrataErro
   
   Screen.MousePointer = vbHourglass
   
   bResult = False
   If Not mvarSys.xDb.Conectado Then GoTo Saida
   If Not VerificaEvento(pEvent, True) Then GoTo Saida
    
   Set TbEvt = CriarObjeto("BANCO_3R.TB_OEVENTOAGENDA", False)
   Set TbEvt.xDb = mvarSys.xDb
   Call TbEvt.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEVENTO:=pEvent.Id)
   Set TbServ = CriarObjeto("BANCO_3R.TB_OSERVICOEVT", False)
   
   Set oQueries = New Collection
   oQueries.Add "Delete From OTAREFAEVT Where IDLOJA=" & TbEvt.IDLOJA & " And IDEVENTO=" & pEvent.Id
   oQueries.Add TbServ.QryDelete(Ch_IDEVENTO:=pEvent.Id, Ch_IDLOJA:=mvarIDLOJA)
   oQueries.Add TbEvt.QryDelete(Ch_IDEVENTO:=pEvent.Id, Ch_IDLOJA:=mvarIDLOJA)
   oQueries.Add QryRemarcado(False, TbEvt.IDLOJA, TbEvt.IDCLIENTE, TbEvt.StartDateTime)
   
   If mvarSys.xDb.Executa(oQueries) Then
      bResult = True
      'Call mvarCalendar_DoReadEvent(pEvent.Id, pEvent)
   Else
      bResult = False
   End If
   GoTo Saida
TrataErro:
   Call ShowError("Sem exclusão em Banco")
Saida:
    Screen.MousePointer = nMouse
End Sub
Private Sub mvarCalendar_DoDeleteRPattern(ByVal pPattern As XtremeCalendarControl.CalendarRecurrencePattern, bResult As Boolean)
   Dim MyTb    As Object 'TB_OEVENTOREC
   Dim nMouse  As Integer

   On Error GoTo TrataErro
   
   nMouse = Screen.MousePointer
   bResult = False
   Screen.MousePointer = vbHourglass
   If Not mvarSys.xDb.Conectado Then GoTo Saida
    
   Set MyTb = GetNewTbEventR
   If MyTb.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEVENTOREC:=pPattern.Id) Then
      bResult = MyTb.Excluir
   End If
   
   GoTo Saida
TrataErro:
   Call ShowError("Sem exclusão em Banco")
Saida:
    Screen.MousePointer = nMouse
End Sub
Private Sub mvarCalendar_DoGetAllEventsRaw(ByVal Events As XtremeCalendarControl.CalendarEvents)
   If Not mvarSys.xDb.Conectado Then
        Exit Sub
    End If
End Sub
Private Sub mvarCalendar_DoGetUpcomingEvents(ByVal dtFrom As Date, ByVal PeriodMinutes As Long, ByVal pEvents As XtremeCalendarControl.CalendarEvents)
   Dim pEvent  As CalendarEvent
   Dim MyRS    As Object 'ADODB.Recordset
   Dim Sql     As String
   Dim nMouse  As Integer

   On Error GoTo TrataErro
   
   nMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   If Not mvarSys.xDb.Conectado Then GoTo Saida
    
   Sql = GetSqlEventoAlerta(dtFrom, PeriodMinutes)
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      While Not MyRS.EOF
         Call MRecordsetParaEvento(MyRS, pEvent, False)
         If Not pEvent Is Nothing Then
            pEvents.Add pEvent
         End If
         MyRS.MoveNext
      Wend
   End If
   GoTo Saida
TrataErro:
   Call ShowError("Sem atualização em Banco")
Saida:
    Screen.MousePointer = nMouse
End Sub
Private Sub mvarCalendar_DoReadEvent(ByVal EventID As Long, pEvent As XtremeCalendarControl.CalendarEvent)
   Dim MyTbEvt As Object ' TB_OEVENTOAGENDA
   Dim nMouse  As Integer
   Dim Sql As String
   Dim MyRS As Object
   
   On Error GoTo TrataErro

   nMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   If Not mvarSys.xDb.Conectado Then GoTo Saida
    
   Sql = Sql_OEVENTOAGENDA(mvarIDLOJA, EventID)
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      '* Popula Atributos do Evento
      Call MRecordsetParaEvento(MyRS, pEvent, False)
      Call MPopulaAtributos(MyRS)
   End If
   GoTo Saida
TrataErro:
   Call ShowError("Sem atualização em Banco")
Saida:
   Screen.MousePointer = nMouse
End Sub
Private Sub mvarCalendar_DoReadRPattern(ByVal PatternID As Long, pPattern As XtremeCalendarControl.CalendarRecurrencePattern)
   Dim MyRS As Object
   Dim MyTb As Object 'TB_OEVENTOREC
   Dim Sql  As String
   Dim nMouse     As Integer

   On Error GoTo TrataErro
   
   nMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   If Not mvarSys.xDb.Conectado Then GoTo Saida
    
   Set MyTb = CriarObjeto("BANCO_3R.TB_OEVENTOREC")
   Sql = MyTb.QrySelect(Ch_IDEVENTOREC:=PatternID)
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      Call MRecordsetParaEventoRec(MyRS, pPattern)
   End If
   
   GoTo Saida
TrataErro:
   Call ShowError("Sem atualização em Banco")
Saida:
    Screen.MousePointer = nMouse
End Sub
Private Sub mvarCalendar_DoRetrieveDayEvents(ByVal dtDay As Date, ByVal pEvents As XtremeCalendarControl.CalendarEvents)
   Dim MyRS    As Object
   Dim RsAux   As Object
   Dim MyEvent As CalendarEvent
   Dim Sql     As String
   Dim nMouse  As Integer
   Dim bExiste As Boolean
   Dim sKey    As String
   Dim bPorDia As Boolean
   
   On Error GoTo TrataErro
   
   nMouse = Screen.MousePointer
   'Screen.MousePointer = vbHourglass
   If Not mvarSys.xDb.Conectado Then GoTo Saida
      
   bPorDia = True
   'bPorDia = mvarCalendar.Visible
   If bPorDia Then
      Sql = GetSqlEventoDia(dtDay, True)
      If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
         While Not MyRS.EOF
            Call MRecordsetParaEvento(MyRS, MyEvent, False)
            If Not MyEvent Is Nothing Then
               pEvents.Add MyEvent
            End If
            MyRS.MoveNext
         Wend
      End If
   
   Else
      If CollMes Is Nothing Then
         Set CollMes = New Collection
      End If
      sKey = "k" & Month(dtDay) & Year(dtDay)
      If Not ExisteItem(CollMes, sKey) Then
         Sql = GetSqlEventoDia(dtDay)
         Call mvarSys.xDb.AbreTabela(Sql, RsAux)
         CollMes.Add RsAux, sKey
         Set RsAux = Nothing
      End If
      Set MyRS = CollMes(sKey)
      MyRS.Filter = ""
      MyRS.Filter = "DIA=" & Day(dtDay)
      While Not MyRS.EOF
         Call MRecordsetParaEvento(MyRS, MyEvent, False)
         If Not MyEvent Is Nothing Then
            pEvents.Add MyEvent
         End If
         MyRS.MoveNext
      Wend
      MyRS.Filter = ""
      Set MyRS = Nothing
   End If
   GoTo Saida
TrataErro:
   Call ShowError("Sem atualização em Banco")
Saida:
    Screen.MousePointer = nMouse
End Sub
Private Sub mvarCalendar_DoUpdateEvent(ByVal pEvent As XtremeCalendarControl.CalendarEvent, bResult As Boolean)
   Dim oTbEvento As Object 'TB_OEVENTOAGENDA
   Dim bSalvou As Boolean
   Dim bMsg    As Boolean
   Dim bisDirt As Boolean
   Dim TBAtend As Object 'TB_OATENDIMENTO
   Dim nID     As Long
   Dim nIDEVT  As Long
   Dim lEvent  As CalendarEvent
   Dim nMouse  As Integer
   Dim Sql     As String
   
   On Error Resume Next
   bMsg = True
   'Set lEvent = mvarCalendar.DataProvider.GetEvent(pEvent.Id)
   If pEvent.CustomProperties("EXIBEMSG") = False Or pEvent.CustomProperties("EXIBIRVENDA") = True Then
      bMsg = False
   Else
      bMsg = True
   End If
   
   On Error GoTo TrataErro
   nMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
      
   nIDEVT = pEvent.Id
   pEvent.CustomProperties("IDSALA") = (pEvent.ScheduleID Mod (mvarIDLOJA * 1000))
   
   bResult = False
   If Not mvarSys.xDb.Conectado Then GoTo Saida
   
   If Not VerificaEvento(pEvent, False, bMsg) Then
      GoTo Saida
   End If
   
   Set oTbEvento = Nothing
'   Set oTbEvento = GetNewTbEvent
   Set oTbEvento = CriarObjeto("BANCO_3R.TB_OEVENTOAGENDA", False)
   Set oTbEvento.xDb = Sys.xDb
   oTbEvento.IDAGENDA = mvarIDAGENDA
   oTbEvento.IDLOJA = mvarIDLOJA
   
   bisDirt = False
   
   If oTbEvento.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEVENTO:=nIDEVT) Then
      'mvarIDCLIENTE = oTbEvento.IDCLIENTE
      'If mvarIDCLIENTE = 0 And oTbEvento.isMeeting = 1 Then
      '   mvarIDCLIENTE = MSalvarCliente
      If xVal(oFrmEvt.TxtIDCLIENTE.Text) = 0 And oTbEvento.ISMEETING = 1 And oTbEvento.IDCLIENTE = 0 Then
         oFrmEvt.TxtIDCLIENTE.Text = MSalvarCliente
      Else
         If oFrmEvt.Visible Then
            oTbEvento.IDCLIENTE = xVal(oFrmEvt.TxtIDCLIENTE.Text)
         Else
            oFrmEvt.TxtIDCLIENTE.Text = oTbEvento.IDCLIENTE
         End If
      End If
      oFrmEvt.TxtIDCLIENTE.Text = oTbEvento.IDCLIENTE
      If oTbEvento.FLGCANCELADO Then
         If VerificarRemarcacao(oTbEvento.IDLOJA, oTbEvento.IDEVENTO) Then
            oFrmEvt.ChkFLGREMARCADO.Value = 1
         Else
            oFrmEvt.ChkFLGREMARCADO.Value = 0
         End If
      Else
         Call mvarSys.xDb.Executa(QryRemarcado(True, oTbEvento.IDLOJA, oTbEvento.IDCLIENTE, oTbEvento.StartDateTime))
      End If
      
      Call MEventoParaClasse(pEvent, oTbEvento)
      If oFrmEvt.isLoad Then
         Call MPopulaAtributos(oTbEvento, True)
      End If
      
      If oTbEvento.isDirt Then
         bisDirt = True
         bResult = MSalvarEvento(oTbEvento)
         
         nID = ExisteAtendimento(pEvent.Id)
         If nID <> 0 Then
            Set TBAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
            Set TBAtend.xDb = mvarSys.xDb
            If TBAtend.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=nID) Then
               TBAtend.dtatend = oTbEvento.StartDateTime
               TBAtend.HHINI = oTbEvento.StartDateTime
               TBAtend.HHFIM = oTbEvento.EndDateTime
               If TBAtend.IDSALA <> oTbEvento.IDSALA Then
                  TBAtend.IDSALA = oTbEvento.IDSALA
                  
                  Dim TbSlMaq As Object
                  Set TbSlMaq = CriarObjeto("BANCO_3R.TB_OSALA_MAQUINA")
                  With TbSlMaq
                     Set .xDb = mvarSys.xDb
                     If .Pesquisar(Ch_IDLOJA:=TBAtend.IDLOJA, Ch_IDSALA:=TBAtend.IDSALA) Then
                        TBAtend.IDMAQUINA = .IDMAQUINA
                     End If
                  End With
                  Set TbSlMaq = Nothing
               End If
                                 
               If TBAtend.isDirt Then
                  bResult = TBAtend.Salvar
               End If
            End If
         End If
      Else
         bResult = True
      End If
           
      bSalvou = bResult And MSalvarSessao(oTbEvento)
      If bSalvou Then
         If oFrmEvt.Visible Then
            Call MSalvarCliente
         End If
         'If bisDirt Then
         '   Call mvarCalendar_DoReadEvent(pEvent.Id, pEvent)
         'End If
      Else
         Call ExibirInformacao("Erro ao gravar evento. [DoUpdate]")
      End If
   End If
   
   GoTo Saida
TrataErro:
   Call ShowError("Sem atualização em Banco")
Saida:
    Screen.MousePointer = nMouse
End Sub
Private Function QryRemarcado(pOk As Boolean, pIDLOJA As Integer, pIDCLIENTE As Integer, pDataIni As String) As String
   Dim Sql As String

   Sql = "Update OEVENTOAGENDA"
   Sql = Sql & " Set FLGREMARCADO=" & IIf(pOk, 1, 0)
   Sql = Sql & " Where FLGCANCELADO=1"
   Sql = Sql & " And StartDateTime<" & SqlDate(pDataIni)
   Sql = Sql & " And IDLOJA=" & pIDLOJA
   Sql = Sql & " And IDCLIENTE=" & pIDCLIENTE
   
   QryRemarcado = Sql
End Function
Private Sub mvarCalendar_DoUpdateRPattern(ByVal pPattern As XtremeCalendarControl.CalendarRecurrencePattern, bResult As Boolean)
   Dim oTbEventoRec As Object 'TB_OEVENTOREC
   Dim nMouse     As Integer

   On Error GoTo TrataErro
   
   nMouse = Screen.MousePointer
   bResult = False
   Screen.MousePointer = vbHourglass
   If Not mvarSys.xDb.Conectado Then GoTo Saida

   Set oTbEventoRec = GetNewTbEventR
   If oTbEventoRec.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEVENTOREC:=pPattern.Id) Then
      'mvarIDCLIENTE = oTbEventoRec.IDCLIENTE
      Call MEventoRecParaClasse(pPattern, oTbEventoRec)
      bResult = oTbEventoRec.Salvar
      If Not bResult Then
         Call ExibirInformacao("Erro ao gravar evento. [DoUpdateR]")
      End If
   End If
   
   GoTo Saida
TrataErro:
   Call ShowError("Sem atualização em Banco")
Saida:
    Screen.MousePointer = nMouse
End Sub
Private Function GetNewTbEventR() As Object 'TB_OEVENTOREC
   Dim oTbEventoR As Object 'TB_OEVENTOREC
      
   Set oTbEventoR = CriarObjeto("BANCO_3R.TB_OEVENTOREC")
   Set oTbEventoR.xDb = Sys.xDb
   oTbEventoR.IDAGENDA = mvarIDAGENDA
   oTbEventoR.IDLOJA = mvarIDLOJA
   
   Set GetNewTbEventR = oTbEventoR
End Function
Private Sub MEventoRecParaClasse(pPattern As XtremeCalendarControl.CalendarRecurrencePattern, ByRef pClasse As Object)  'TB_OEVENTOREC)
   On Error GoTo TrataErro

    pClasse.MasterEventId = pPattern.MasterEventId
    
    pClasse.EventStartTime = pPattern.StartTime
    pClasse.EventDuration = pPattern.DurationMinutes
    
    pClasse.PatternStartDate = pPattern.StartDate
    
    pClasse.PatternEndMethod = pPattern.EndMethod
    pClasse.PatternEndDate = pPattern.EndDate
    pClasse.PatternEndAfterOccurrences = pPattern.EndAfterOccurrences
    
    pClasse.OptionsData1 = pPattern.Options.Data1
    pClasse.OptionsData2 = pPattern.Options.Data2
    pClasse.OptionsData3 = pPattern.Options.Data3
    pClasse.OptionsData4 = pPattern.Options.Data4
    
    pClasse.CustomPropertiesXMLData = pPattern.CustomProperties.SaveToString

   Exit Sub
TrataErro:
   Call ShowError("Sem atualização de Classe")
End Sub
Private Function GetNewTbEvent() As Object 'TB_OEVENTOAGENDA
   Dim oTbEvento As Object 'TB_OEVENTOAGENDA
      
   Set oTbEvento = Nothing
   Set oTbEvento = CriarObjeto("BANCO_3R.TB_OEVENTOAGENDA")
   Set oTbEvento.xDb = Sys.xDb
   oTbEvento.IDAGENDA = mvarIDAGENDA
   oTbEvento.IDLOJA = mvarIDLOJA
   
   Set GetNewTbEvent = oTbEvento
End Function
Private Sub MEventoParaClasse(pEvento As CalendarEvent, ByRef pClasse As Object) ' TB_OEVENTOAGENDA)
   On Error GoTo TrataErro
      
   '***************
   '* Propriedades Customizadas
   If oFrmEvt.isLoad Then
      'If pClasse.IDCLIENTE <> mvarIDCLIENTE Then
      '   MsgBox "XXX Cliente Diferente"
      'End If
      pEvento.CustomProperties("IDCLIENTE") = xVal(oFrmEvt.TxtIDCLIENTE.Text)    'mvarIDCLIENTE
      If oFrmEvt.CmbSchedule.ListCount = 0 Or oFrmEvt.CmbSchedule.ListIndex = -1 Then
         pEvento.CustomProperties("IDSALA") = (pEvento.ScheduleID Mod mvarIDLOJA * 1000)
      Else
         pEvento.CustomProperties("IDSALA") = xVal(oFrmEvt.CmbSchedule.ItemData(oFrmEvt.CmbSchedule.ListIndex))
      End If
      If oFrmEvt.Visible Then
         pEvento.CustomProperties("FLGCONFIRMADO") = oFrmEvt.ChkFLGCONFIRMADO.Value
         pEvento.CustomProperties("FLGCANCELADO") = oFrmEvt.ChkFLGCANCELADO.Value
         pEvento.CustomProperties("FLGREMARCADO") = oFrmEvt.ChkFLGREMARCADO.Value
         pEvento.CustomProperties("FLGAVALIACAO") = oFrmEvt.ChkFLGAVALIACAO.Value
      End If
   Else
      'mvarIDCLIENTE = pClasse.IDCLIENTE
      pEvento.CustomProperties("IDCLIENTE") = pClasse.IDCLIENTE
      pEvento.CustomProperties("IDSALA") = pClasse.IDSALA
      pEvento.CustomProperties("FLGCONFIRMADO") = pClasse.FLGCONFIRMADO
      pEvento.CustomProperties("FLGCANCELADO") = pClasse.FLGCANCELADO
      pEvento.CustomProperties("FLGREMARCADO") = pClasse.FLGREMARCADO
      pEvento.CustomProperties("FLGAVALIACAO") = pClasse.FLGAVALIACAO
   End If
   Call SetCustomIcons(pEvento)
   pClasse.IDCLIENTE = xVal(pEvento.CustomProperties("IDCLIENTE")) 'mvarIDCLIENTE
   pClasse.IDSALA = pEvento.CustomProperties("IDSALA")
   pClasse.FLGCONFIRMADO = pEvento.CustomProperties("FLGCONFIRMADO")
   pClasse.FLGCANCELADO = pEvento.CustomProperties("FLGCANCELADO")
   pClasse.FLGREMARCADO = pEvento.CustomProperties("FLGREMARCADO")
   pClasse.FLGAVALIACAO = pEvento.CustomProperties("FLGAVALIACAO")
   
   
   '* Fim de Propriedades Customizadas
   '***************
   
   pClasse.Subject = pEvento.Subject
   pClasse.Location = pEvento.Location
   pClasse.Body = pEvento.Body
   
   pClasse.ISMEETING = IIf(pEvento.MeetingFlag, 1, 0)
   pClasse.isPrivate = IIf(pEvento.PrivateFlag, 1, 0)
   
   pClasse.LabelID = pEvento.Label
   pClasse.BusyStatus = pEvento.BusyStatus
   pClasse.ImportanceLevel = pEvento.Importance
   
   pClasse.StartDateTime = Format(pEvento.StartTime, "dd/mm/yyyy hh:mm:00")
   pClasse.EndDateTime = Format(pEvento.EndTime, "dd/mm/yyyy hh:mm:00")
       
   pClasse.isAllDayEvent = IIf(pEvento.AllDayEvent, 1, 0)
   
   If pClasse.isAllDayEvent = 1 Then
      pClasse.EndDateTime = pClasse.StartDateTime
   End If
   
   pClasse.isReminder = IIf(pEvento.Reminder, 1, 0)
   pClasse.ReminderMinutesBeforeStart = pEvento.ReminderMinutesBeforeStart
   pClasse.RemainderSoundFile = pEvento.ReminderSoundFile
           
   pClasse.RecurrenceState = pEvento.RecurrenceState
   
   If pEvento.RecurrenceState = xtpCalendarRecurrenceMaster Or _
      pEvento.RecurrenceState = xtpCalendarRecurrenceException Then
      pClasse.IDEVENTOREC = pEvento.RecurrencePattern.Id
   Else
      pClasse.IDEVENTOREC = 0
   End If
   
   If pEvento.RExceptionStartTimeOrig <> "00:00:00" Then
      pClasse.RExceptionStartTimeOrig = pEvento.RExceptionStartTimeOrig
   End If
   If pEvento.RExceptionEndTimeOrig <> "00:00:00" Then
      pClasse.RExceptionEndTimeOrig = pEvento.RExceptionEndTimeOrig
   End If
   pClasse.IsRecurrenceExceptionDeleted = IIf(pEvento.RExceptionDeleted, 1, 0)
       
   pClasse.CustomIconsIDs = pEvento.CustomIcons.SaveToString
           
   If pClasse.ScheduleID <> pEvento.ScheduleID Then
      'mvarIDAGENDA = 1
      pClasse.ScheduleID = pEvento.ScheduleID
      pClasse.IDAGENDA = mvarIDAGENDA
      pClasse.IDSALA = pEvento.CustomProperties("IDSALA")
   End If
   
   If Not pClasse.isDirt Then
      pClasse.CustomPropertiesXMLData = pEvento.CustomProperties.SaveToString
      pClasse.Salvar
   Else
      pClasse.CustomPropertiesXMLData = pEvento.CustomProperties.SaveToString
   End If
   
   Exit Sub
TrataErro:
   Call ShowError("Sem atualização de Classe")
End Sub
Private Sub MRecordsetParaEvento(ByVal pRs As Object, ByRef pEvent As CalendarEvent, ByVal bRException) 'As CalendarEvent
   Dim nEventID  As Long
   Dim sTag      As Long
   
   On Error GoTo TrataErro
    
   Set pEvent = Nothing
       
   If pRs.EOF Then Exit Sub
       
   nEventID = pRs("IDEVENTO")
   Set pEvent = mvarCalendar.DataProvider.CreateEventEx(nEventID)
   If pEvent Is Nothing Then Exit Sub
       
   pEvent.Subject = pRs("Subject")
   pEvent.Location = pRs("Location")
   pEvent.Body = pRs("Body")
   pEvent.MeetingFlag = IIf(pRs("IsMeeting") <> 0, True, False)
   pEvent.PrivateFlag = IIf(pRs("IsPrivate") <> 0, True, False)
   
   pEvent.Label = pRs("LabelID")
   If IsNull(pRs("IDCLIENTE")) Then
      pEvent.BusyStatus = pRs("BusyStatus")
   Else
      If pRs("IDCLIENTE") = 2156 Then
         pEvent.BusyStatus = ShowBusyStatus(pRs("IDCLIENTE"), pRs("BusyStatus"))
      Else
         pEvent.BusyStatus = ShowBusyStatus(pRs("IDCLIENTE"), pRs("BusyStatus"))
      End If
   End If
   pEvent.Importance = pRs("ImportanceLevel")
   
   pEvent.StartTime = pRs("StartDateTime")
   pEvent.EndTime = pRs("EndDateTime")
       
   pEvent.AllDayEvent = IIf(pRs("IsAllDayEvent") <> 0, True, False)
   
   pEvent.Reminder = IIf(pRs("IsReminder") <> 0, True, False)
   pEvent.ReminderMinutesBeforeStart = pRs("ReminderMinutesBeforeStart")
   pEvent.ReminderSoundFile = pRs("RemainderSoundFile")
   
   pEvent.CustomProperties.LoadFromString pRs("CustomPropertiesXMLData")
   pEvent.CustomIcons.LoadFromString pRs("CustomIconsIDs")
      
   If pRs("ScheduleID") = 0 Then
      pEvent.ScheduleID = mvarIDLOJA * 1000 + 1
   Else
      pEvent.ScheduleID = pRs("ScheduleID")
   End If
   
   If bRException Then
      pEvent.MakeAsRException
      
      pEvent.RExceptionStartTimeOrig = pRs("RExceptionStartTimeOrig")
      pEvent.RExceptionEndTimeOrig = pRs("RExceptionEndTimeOrig")
      pEvent.RExceptionDeleted = IIf(pRs("ISRecurrenceExceptionDeleted") <> 0, True, False)
   End If
    
   If Not bRException Then
      ' "process_RecurrenceState" and "process_RecurrencePatternID" properties
      ' are used to process master events.
      '
      ' If they are set and RecurrenceStaie is Master Data provider will
      ' fier DoReadRPattern event and make event as Master.
      ' And it will also generate ocurrences for RetrieveDayEvents method.
      '
      ' Thise properties are temporary and they will be removed by data provider.
      '
      ' If these properties are not set data provider expect that master event
      ' is already compleated - CreateRecurrence method is called and
      ' Recurrence pattern is set.
      '
      ' This mechanism is usefull for DB data providers, when events and patterns
      ' are stored separately (in tables).
      ' But if events stored in some memory collection or array
      ' it should not be used because master event store recurrence pattern inside.
      '
      pEvent.CustomProperties("process_RecurrenceState") = pRs("RecurrenceState").Value
      pEvent.CustomProperties("process_RecurrencePatternID") = pRs("IDEVENTOREC").Value
   End If
   
   '* Propriedades Customizadas
   Call MCustomizacaoParaEvento(pRs, pEvent, bRException)
    
   Exit Sub
TrataErro:
   If Err.Number <> 0 Then
      ShowError "Erro ao ler evento."
      Resume Next
   End If
End Sub
Private Function ShowBusyStatus(pIDCLIENTE As Integer, pBusyStatus As Integer) As Integer
   Dim Sql As String
   Dim MyRS As Object
   
   Sql = ""
   Sql = Sql & "Select Top 1 Cast(( IsNull((SELECT SUM(I.QTDVENDA)"
   Sql = Sql & " FROM CITENSVENDA I"
   Sql = Sql & " JOIN SPRODUTO P ON I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD And P.ESERVICO=1 And P.EVENDA=1"
   Sql = Sql & " Where i.IDLOJA = V.IDLOJA"
   Sql = Sql & "  And I.IDVENDA=V.IDVENDA),0)"
   Sql = Sql & "-"
   Sql = Sql & "  IsNull((SELECT COUNT(S.IDSESSAO)"
   Sql = Sql & "  From OSESSAO S"
   Sql = Sql & "  Where S.IDVENDA = V.IDVENDA"
   Sql = Sql & "  And isNull(IDITEM,0)<>0"
   Sql = Sql & "  And S.IDTPSERVICO<>1),0)"
   Sql = Sql & ") as Integer) [Saldo]"
   Sql = Sql & " From CVENDA V"
   Sql = Sql & " Where V.IDLOJA =" & mvarIDLOJA
   Sql = Sql & " And V.IDCLIENTE=" & pIDCLIENTE
   Sql = Sql & " Order By 1 Desc"
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      ShowBusyStatus = IIf(MyRS("Saldo") <> 0, 0, pBusyStatus)
   Else
      ShowBusyStatus = pBusyStatus
   End If
End Function

Private Sub MCustomizacaoParaEvento(ByVal pRs As Object, ByRef pEvent As CalendarEvent, ByVal bRException) 'As CalendarEvent
   
   Call SetEventFlag(pEvent, pRs)
   
   Exit Sub
TrataErro:
   If Err.Number <> 0 Then
      ShowError "Erro ao ler evento."
   End If
End Sub

Private Sub MRecordsetParaEventoRec(pRs As Object, ByRef pRPattern As CalendarRecurrencePattern)
   Dim nPatternID  As Long
   
   On Error GoTo TrataErro
    
   Set pRPattern = Nothing
       
   If pRs.EOF Then Exit Sub
       
   nPatternID = pRs("IDEVENTOREC")
   Set pRPattern = mvarCalendar.DataProvider.CreateRecurrencePattern(nPatternID)
   If pRPattern Is Nothing Then Exit Sub
            
   pRPattern.MasterEventId = pRs("MasterEventID")
   
   pRPattern.StartTime = pRs("EventStartTime")
   pRPattern.DurationMinutes = pRs("EventDuration")
   
   pRPattern.StartDate = pRs("PatternStartDate")
   
   pRPattern.EndMethod = pRs("PatternEndMethod")
   
   If pRPattern.EndMethod = xtpCalendarPatternEndDate Then
       pRPattern.EndDate = pRs("PatternEndDate")
   ElseIf pRPattern.EndMethod = xtpCalendarPatternEndAfterOccurrences Then
       pRPattern.EndAfterOccurrences = pRs("PatternEndAfterOccurrences")
   Else
       Debug.Assert pRPattern.EndMethod = xtpCalendarPatternEndNoDate
   End If

   pRPattern.Options.Data1 = pRs("OptionsData1")
   pRPattern.Options.Data2 = pRs("OptionsData2")
   pRPattern.Options.Data3 = pRs("OptionsData3")
   pRPattern.Options.Data4 = pRs("OptionsData4")
   
   pRPattern.CustomProperties.LoadFromString pRs("CustomPropertiesXMLData")
       
   MReadRPatternExceptions pRPattern
   
   Exit Sub
TrataErro:
   If Err.Number <> 0 Then
      ShowError "Erro ao ler evento recorrente."
   End If
End Sub
Public Sub MPopulaAtributos(pTbEvento As Object, Optional isUpdate As Boolean)   'TB_OEVENTOAGENDA)
   Dim Sql As String
   Dim nIDLOJA As Long
   Dim nIDEVENTO As Long
   Dim nIDCLIENTE As Long
   Dim nFLGCONFIRMADO As Integer
   Dim nFLGCANCELADO As Integer
   Dim nFLGREMARCADO As Integer
   
   Static sIDEVENTO As Long
   Static sQtd As Integer
   Dim bRep As Boolean
   
   On Error Resume Next
   If Not oFrmEvt.isLoad Then Exit Sub
   
   If TypeName(pTbEvento) = "Recordset" Then
      nIDLOJA = pTbEvento("IDLOJA")
      nIDEVENTO = pTbEvento("IDEVENTO")
      nIDCLIENTE = pTbEvento("IDCLIENTE")
      nFLGCONFIRMADO = pTbEvento("FLGCONFIRMADO")
      nFLGCANCELADO = pTbEvento("FLGCANCELADO")
      nFLGREMARCADO = pTbEvento("FLGREMARCADO")
   Else
      nIDLOJA = pTbEvento.IDLOJA
      nIDEVENTO = pTbEvento.IDEVENTO
      nIDCLIENTE = pTbEvento.IDCLIENTE
      nFLGCONFIRMADO = pTbEvento.FLGCONFIRMADO
      nFLGCANCELADO = pTbEvento.FLGCANCELADO
      nFLGREMARCADO = pTbEvento.FLGREMARCADO
   End If
   bRep = (sIDEVENTO = nIDEVENTO)
   sQtd = IIf(bRep, sQtd + 1, 1)
   sQtd = IIf(sQtd >= 7, 1, sQtd)
   If bRep And sQtd >= 6 Then
      Exit Sub
   End If
   
   sIDEVENTO = nIDEVENTO
   
   oFrmEvt.ChkFLGCONFIRMADO.Value = nFLGCONFIRMADO
   oFrmEvt.ChkFLGCANCELADO.Value = nFLGCANCELADO
   oFrmEvt.ChkFLGREMARCADO.Value = nFLGREMARCADO
   oFrmEvt.ChkFLGCANCELADO.Visible = oFrmEvt.ChkFLGCANCELADO.Value
   oFrmEvt.ChkFLGREMARCADO.Visible = oFrmEvt.ChkFLGCANCELADO.Value
   
   '*********
   '* Popula Remarcação
   If oFrmEvt.ChkFLGREMARCADO.Value Or nIDCLIENTE <> 0 Then
      Sql = "Select StartDateTime, FLGCANCELADO" & vbNewLine
      Sql = Sql & " From OEVENTOAGENDA" & vbNewLine
      Sql = Sql & " Where IDLOJA=" & nIDLOJA & vbNewLine
      Sql = Sql & " And IDCLIENTE=" & nIDCLIENTE & vbNewLine
      Sql = Sql & " And StartDateTime >" & SqlDate(pTbEvento.StartDateTime) & vbNewLine
      Sql = Sql & " And FLGCANCELADO=0" & vbNewLine
      Sql = Sql & " Order By StartDateTime"
      If mvarSys.xDb.AbreTabela(Sql) Then
         oFrmEvt.ChkFLGREMARCADO.Caption = "Remarcado ->" & Format(mvarSys.xDb.RsAux("StartDateTime"), "DD/MM HH:MM")
      End If
   End If
   
   '*********
   '* Popula Cliente
   If mvarTlEvt.isLoad Then
      If xVal(oFrmEvt.TxtIDCLIENTE.Text) = 0 Then
         oFrmEvt.TxtIDCLIENTE.Text = nIDCLIENTE
      End If
      mvarTlEvt.CalEvent.CustomProperties("IDCLIENTE") = nIDCLIENTE
   End If
'   If mvarTlEvt.isLoad Then
'      If xVal(mvarTlEvt.CalEvent.CustomProperties("IDCLIENTE")) <> 0 Then
'         nIDCLIENTE = xVal(mvarTlEvt.CalEvent.CustomProperties("IDCLIENTE"))
'         oFrmEvt.TxtIDCLIENTE.Text = nIDCLIENTE
'      End If
'   End If
   If mvarTlEvt.isLoad Then
      If nIDCLIENTE <> xVal(oFrmEvt.TxtIDCLIENTE.Text) Then
         Call MPopulaCliente(nIDCLIENTE)
      End If
   End If
   
   '*********
   '* Popula ServicoEvt
   'If Not oFrmEvt.Visible Then
   If Not isUpdate Then
      Call MPopulaServicoEVT(nIDEVENTO)
   End If
   
   '*********
   '* Popula Tarefas
   If mvarTlEvt.isLoad Then
      Call MPopulaTarefas(nIDEVENTO)
   End If
End Sub
Private Sub MPopulaCliente(pIDCLIENTE As Long)
   Dim Sql As String
   Dim MyRS As Object
   
   Sql = Sql_OCLIENTE(mvarIDLOJA, pIDCLIENTE)
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      oFrmEvt.TxtIDCLIENTE.Text = MyRS("IDCLIENTE")
      oFrmEvt.TxtNOME.Text = MyRS("NOME")
      oFrmEvt.TxtTEL1.Text = MyRS("TEL1") & ""
      oFrmEvt.TxtTEL2.Text = MyRS("TEL2") & ""
      oFrmEvt.TxtEMAIL.Text = MyRS("EMAIL") & ""
      If MyRS("OBS") <> "" Then
         oFrmEvt.CmdIDCLIENTE.Caption = "Cliente *"
      End If
   End If
End Sub
Private Sub MPopulaTarefas(pEventID As Long)
   Dim MyRS As Object
   Dim i As Integer
   Dim Sql As String
   Dim vCol As String
   
   With oFrmEvt.GrdTarefas
      .Clear
      .Editable = False
      .BorderStyle = igBorderThin
      .RowMode = True
      .ImageList = oFrmEvt.lstImage
      
      '.PaintManager.ColumnStyle = xtpColumnOffice2007
      '.PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
      '.PaintManager.RefreshMetrics
         
      '.SetImageList oFrmEvt.ImageList1
      '.SetCustomDraw xtpCustomBeforeDrawRow
      '.MultipleSelection = False
      '.AllowColumnRemove = False
      '.AllowColumnReorder = False
      '.AllowColumnResize = False
      '.AllowColumnSort = False
      '.AllowEdit = False

      Sql = "Select ' ' [Sit.], Cast(Convert(Varchar(10), T.DTTAREFA, 103) As DateTime) [Data], P.DSCTPTAREFA [Tarefa]"
      Sql = Sql & ", T.SITTAREFA, T.FLGDELETE, T.DSCTAREFA, P.DSCTPTAREFA"
      Sql = Sql & ", T.IDLOJA, T.IDTAREFA, T.IDCLIENTE"
      Sql = Sql & " From OTAREFAEVT T"
      Sql = Sql & " Left Join OTIPOTAREFA P On T.IDTPTAREFA=P.IDTPTAREFA And T.IDLOJA=P.IDLOJA"
      Sql = Sql & " Where T.IDLOJA = " & mvarIDLOJA
      Sql = Sql & " And T.IDEVENTO=" & pEventID
      Sql = Sql & " And T.FLGDELETE=0"
      Sql = Sql & " Order By 1"
      
      If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
         .FillFromRS MyRS
         
         Call EscondeCampos(oFrmEvt.GrdTarefas, MyRS, Array("Sit.", "Data", "Tarefa"))
         For i = 0 To MyRS.Fields.Count - 1
            vCol = MyRS.Fields(i).Name
            .ColVisible(vCol) = InArray(vCol, Array("Sit.", "Data", "Tarefa"))
         Next
          
         .ColWidth("Sit.") = 30
         .ColWidth("Data") = 80
         .ColWidth("Tarefa") = 275
         
         For i = 1 To .RowCount
            .RowHeight(i) = 20
            If .CellValue(i, "SITTAREFA") = 5 Then
               .CellExtraIcon(i, 1) = 0
            ElseIf .CellValue(i, "SITTAREFA") = 10 Then
               .CellExtraIcon(i, 1) = 1
            End If
            If .CellValue(i, "FLGDELETE") = 1 Then
               .CellExtraIcon(i, 1) = 2
            End If
         Next
      End If
   End With
End Sub
Private Sub EscondeCampos(Optional Grd, Optional MyRS, Optional pArray)
   Dim i As Integer
   
   If MyRS Is Nothing Then Exit Sub
   If MyRS.state = 0 Then Exit Sub

End Sub

Private Sub MPopulaServicoEVT(pEventID As Long, Optional pVerifAtend = True)
   Dim pLabel  As CalendarEventLabel
   Dim MyTb    As Object 'TB_OSERVICOEVT
   Dim TbEvt   As Object 'TB_OEVENTOAGENDA
   Dim sTag    As String
   Dim i       As Integer
   Dim vCtrl   As Control
   Dim nID     As Long
   Dim bAchou  As Boolean
   Dim RsEvt   As Object
   Dim RsServ  As Object

   If oFrmEvt.GrdSERVICOEVT Is Nothing Then Exit Sub
   
   Set TbEvt = CriarObjeto("BANCO_3R.TB_OEVENTOAGENDA")
   Set TbEvt.xDb = mvarSys.xDb
   If TbEvt.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEVENTO:=pEventID) Then
      Set MyTb = CriarObjeto("BANCO_3R.TB_OSERVICOEVT")
      Set MyTb.xDb = mvarSys.xDb
      If MyTb.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEVENTO:=pEventID) Then
         
         oFrmEvt.GrdSERVICOEVT.Clear
         While Not MyTb.Rs.EOF
            MyTb.Popula
                                 
            mvarTlEvt.IncluiGrdLinha
            With oFrmEvt.GrdSERVICOEVT
               i = .RowCount - 1
               If i = 1 Then
                  oFrmEvt.ChkFLGAVALIACAO.Value = IIf(MyTb.IDTPSERVICO = 1, vbChecked, vbUnchecked)
               End If
               If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
               Call SetTag(sTag, "IDSERVICOEVT", MyTb.IDSERVICOEVT)
               Call SetTag(sTag, "IDTPSERVICO", MyTb.IDTPSERVICO)
               Call SetTag(sTag, "IDTPTRATAMENTO", MyTb.IDTPTRATAMENTO)
               Call SetTag(sTag, "IDAREA", MyTb.IDAREA)
               .RowTag(i) = sTag
               
               .CellValue(i, "IDSERVICOEVT") = MyTb.IDSERVICOEVT
               .CellValue(i, "IDTPSERVICO") = MyTb.IDTPSERVICO
               .CellValue(i, "IDTPTRATAMENTO") = MyTb.Rs("IDTPTRATAMENTO")
               .CellValue(i, "IDAREA") = MyTb.IDAREA
               
               '.CellText(i, "IDTPSERVICO") = .Combos("IDTPSERVICO").ItemText(MyTb.IDTPSERVICO)
               '.CellText(i, "IDTPTRATAMENTO") = .Combos("IDTPTRATAMENTO").ItemText(MyTb.IDTPTRATAMENTO)
               '.CellText(i, "IDAREA") = .Combos("IDAREA").ItemText(MyTb.IDAREA)
               
               
               If i = 1 Then
                  oFrmEvt.CmbShowTimeAs.Enabled = False
                  oFrmEvt.CmbShowTimeAs.ListIndex = DefineShowAs(MyTb.IDTPSERVICO)

                  nID = TbEvt.LabelID '  DefineLabelID(MyTb.IDTPTRATAMENTO, MyTb.IDTPSERVICO, TbEvt.FLGCANCELADO)
                  If Not mvarCalendar Is Nothing Then
                     Set pLabel = mvarCalendar.DataProvider.LabelList.Find(nID)
                     If Not pLabel Is Nothing Then
                        oFrmEvt.ctrlColor.BackColor = pLabel.Color
                     End If
                  End If
               End If
               
            End With
            MyTb.Rs.MoveNext
         Wend
         
         If pVerifAtend Then
            '*****************
            '* Verifica se existe Atendiemnto.
            nID = ExisteAtendimento(MyTb.IDEVENTO)
            Call SetTag(oFrmEvt.GrpSessao, "IDATENDIMENTO", nID)
            If nID <> 0 Then
               oFrmEvt.ChkMeeting.Enabled = False
               
               If CDate(mvarSys.xDb.Sysdate) >= CDate(Format(TbEvt.StartDateTime, "dd/mm/yyyy")) Then
                  oFrmEvt.GrpSessao.Enabled = False
                  For Each vCtrl In oFrmEvt.Controls
                     If vCtrl.Container Is oFrmEvt.GrpSessao Then
                        vCtrl.Enabled = False
                     End If
                  Next
                  
               End If
            End If
         End If
      End If
   End If
   Set TbEvt = Nothing
   Set MyTb = Nothing
End Sub
Public Function ExisteAtendimento(pIDEVENTO As Long) As Long
   Dim Sql     As String
   
   Sql = Sql_OATENDIMENTO(mvarIDLOJA, pIDEVENTO:=pIDEVENTO)
   If mvarSys.xDb.AbreTabela(Sql) Then
      ExisteAtendimento = mvarSys.xDb.RsAux("IDATENDIMENTO")
   End If
End Function
Public Function ExcluiAtendimento(pIDEVENTO As Long) As Long
   Dim TBAtend    As Object 'TB_OATENDIMENTO
   Dim TbProd     As Object 'TB_OATENDIMENTO_PRODUTO
   Dim TbSessao   As Object 'TB_OSESSAO
   Dim TbReacao   As Object 'TB_OSESSAO_REACAOADV
   Dim Queries    As Collection
   
   Set Queries = New Collection
   
   Set TBAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
   Set TBAtend.xDb = mvarSys.xDb
   With TBAtend
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDEVENTO=" & pIDEVENTO) Then
         Set TbSessao = CriarObjeto("BANCO_3R.TB_OSESSAO")
         Set TbSessao.xDb = mvarSys.xDb
         If TbSessao.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=.IDATENDIMENTO) Then
            Set TbReacao = CriarObjeto("BANCO_3R.TB_OSESSAO_REACAOADV")
            Set TbReacao.xDb = mvarSys.xDb
            
            Queries.Add TbReacao.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=.IDATENDIMENTO)
            Queries.Add TbSessao.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=.IDATENDIMENTO)
         End If
         Set TbProd = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_PRODUTO")
         Set TbProd.xDb = mvarSys.xDb
         
         Queries.Add TbProd.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=.IDATENDIMENTO)
         Queries.Add .QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=.IDATENDIMENTO)
      End If
   End With
   
   If Queries.Count > 0 Then
      Call ExibirResultado(mvarSys, mvarSys.xDb.Executa(Queries))
   End If

End Function
Private Function MSalvarCliente() As Long
   Dim TBCliente  As Object
   Dim bClienteOK As Boolean
   Dim bAchou As Boolean
   Dim sMsg As String
   Dim bResult As Boolean
   Dim Sql As String
   
   Set TBCliente = CriarObjeto("BANCO_3R.TB_OCLIENTE")
   With TBCliente
      Set .xDb = mvarSys.xDb
      bAchou = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCLIENTE:=xVal(oFrmEvt.TxtIDCLIENTE.Text))
      If (oFrmEvt.ChkMeeting.Value = vbChecked) Then
         If bAchou Then
            sMsg = vbNewLine & vbNewLine
            If xVal(oFrmEvt.TxtIDCLIENTE) <> .IDCLIENTE Then sMsg = sMsg & "Id.: " & .IDCLIENTE & " -> " & xVal(oFrmEvt.TxtIDCLIENTE.Text) & vbNewLine
            If Trim(oFrmEvt.TxtNOME.Text) <> .NOME Then sMsg = sMsg & "Nome: " & .NOME & " -> " & Trim(oFrmEvt.TxtNOME.Text) & vbNewLine
            If Trim(oFrmEvt.TxtTEL1.Text) <> .TEL1 Then sMsg = sMsg & "Celular: " & .TEL1 & " -> " & Trim(oFrmEvt.TxtTEL1.Text) & vbNewLine
            If Trim(oFrmEvt.TxtTEL2.Text) <> .TEL2 Then sMsg = sMsg & "Telefone: " & .TEL2 & " -> " & Trim(oFrmEvt.TxtTEL2.Text) & vbNewLine
            If Trim(oFrmEvt.TxtEMAIL.Text) <> .EMAIL Then sMsg = sMsg & "e-Mail: " & .EMAIL & " -> " & Trim(oFrmEvt.TxtEMAIL.Text) & vbNewLine
         Else
            .FLGMARKETING = 1
            .FLGAGENDA = 1
            .ATIVO = 1
            
            Sql = "Select * From OCLIENTE Where NOME Like '" & Trim(GetWords(oFrmEvt.TxtNOME.Text)(1)) & "%'"
            If mvarSys.xDb.AbreTabela(Sql) Then
               .SEXO = mvarSys.xDb.RsAux("SEXO")
            End If
         End If

         If Not bAchou Then .IDLOJA = mvarIDLOJA
         .IDLOJA0 = IIf(mvarSys.GetParam("OCLIENTEUNICO", Default:=0) = "1", 0, mvarIDLOJA)
         If .isDirt Then .Alterar
         '.IDCLIENTE = Trim(oFrmEvt.TxtIDCLIENTE.Text)
         .NOME = Trim(oFrmEvt.TxtNOME.Text)
         .TEL1 = Trim(oFrmEvt.TxtTEL1.Text)
         .TEL2 = Trim(oFrmEvt.TxtTEL2.Text)
         .EMAIL = Trim(oFrmEvt.TxtEMAIL.Text)
         
         If .IDCLIENTE = 0 Then .DTCADASTRO = mvarSys.xDb.Sysdate
         
         bClienteOK = (Trim(.NOME) <> "") And (Trim(.TEL1 & .TEL2 & .EMAIL) <> "")
         If bClienteOK Then
            If .isDirt Then
               If .IDCLIENTE = 0 Then
                  bResult = .Salvar
               Else
                  If vbYes = ExibirPergunta("Atualizar dados do Cliente?" & sMsg) Then
                     bResult = .Alterar
                  End If
               End If
            Else
               .ATIVO = 1
               If .isDirt Then
                  Call .Alterar
               End If
            End If
         End If
      End If
      'mvarIDCLIENTE = .IDCLIENTE
      MSalvarCliente = .IDCLIENTE
   End With
End Function
Public Function MSalvarEvento(ByRef pTbEvento As Object) As Boolean
   Dim NgTarefa As Object
   Dim cQueries   As Collection
   Dim bExiste As Boolean
   Dim bTeste As Boolean
   
   On Error Resume Next
   bExiste = True
   '* Salvar Evento
   Set cQueries = New Collection
   If pTbEvento.IDEVENTO = 0 Then
      bExiste = False
      bTeste = (1 = Val(oFrmEvt.GrdSERVICOEVT.CellValue(1, "IDTPSERVICO")))
      If mvarSys.xDb.Executa(pTbEvento.QryInsert) Then ' oTbEvento.Salvar
         Call pTbEvento.Pesquisar(Ch_IDLOJA:=pTbEvento.IDLOJA, Ch_IDEVENTO:=mvarSys.xDb.ADORs("IDEVENTO"))
      End If
   ElseIf pTbEvento.isDirt Then
      cQueries.Add pTbEvento.QrySave
   End If
   
   '* Salvar Tarefas
   Set NgTarefa = CriarObjeto("Tarefa3R.NG_TAREFA", False)
   Set NgTarefa.Sys = mvarSys
   Call NgTarefa.SalvarTarefas(pTbEvento, cQueries, bTeste)
   Set NgTarefa = Nothing
   
   MSalvarEvento = mvarSys.xDb.Executa(cQueries) ' oTbEvento.Salvar

'   If Not mvarCalendar Is Nothing And bExiste Then
'      mvarCalendar.DataProvider.ClearCache
'      mvarCalendar.Populate
'   End If
End Function
Private Function MSalvarSessao(ByRef pTbEvento As Object) As Boolean
   Dim i          As Long
   Dim cQueries   As Collection
   Dim TbServico  As Object 'TB_OSERVICOEVT
         
   MSalvarSessao = True
   
   If Not mvarTlEvt.isLoad Then
      'Call MCopiarSessao(pTbEvento)
      Exit Function
   End If
   If oFrmEvt Is Nothing Then Exit Function
   If Not oFrmEvt.isLoad Then Exit Function
   If Not (oFrmEvt.ChkMeeting.Value = vbChecked) Then
      Call SetTag(oFrmEvt.GrdSERVICOEVT, "ISDIRT", "1")
   End If
         
   Set cQueries = New Collection
     
   'mvarIDCLIENTE = MSalvarCliente
   If pTbEvento.IDCLIENTE <> xVal(oFrmEvt.TxtIDCLIENTE.Text) Or xVal(oFrmEvt.TxtIDCLIENTE.Text) = 0 Then
      oFrmEvt.TxtIDCLIENTE.Text = MSalvarCliente
   End If
   
   pTbEvento.IDCLIENTE = xVal(oFrmEvt.TxtIDCLIENTE.Text)
   pTbEvento.FLGCONFIRMADO = oFrmEvt.ChkFLGCONFIRMADO.Value
   pTbEvento.FLGCANCELADO = oFrmEvt.ChkFLGCANCELADO.Value
   pTbEvento.FLGREMARCADO = oFrmEvt.ChkFLGREMARCADO.Value
   'If pTbEvento.FLGCONFIRMADO = 1 Then
   '   pTbEvento.FLGCANCELADO = 0
   '   oFrmEvt.ChkFLGCANCELADO.Value
   'End If
   
   If pTbEvento.isDirt Then cQueries.Add pTbEvento.QrySave
   
   If GetTag(oFrmEvt.GrdSERVICOEVT, "ISDIRT", 0) = 1 And pTbEvento.IDEVENTO > 0 And oFrmEvt.Visible Then
      
      Set TbServico = CriarObjeto("BANCO_3R.TB_OSERVICOEVT", False)
      Set TbServico.xDb = mvarSys.xDb
      TbServico.IDLOJA = pTbEvento.IDLOJA
      TbServico.IDEVENTO = pTbEvento.IDEVENTO

      cQueries.Add TbServico.QryDelete(Ch_IDLOJA:=TbServico.IDLOJA, Ch_IDEVENTO:=TbServico.IDEVENTO, Ch_Where:="IDSERVICOEVT >= " & oFrmEvt.GrdSERVICOEVT.RowCount)
      If oFrmEvt.ChkMeeting.Value = vbChecked Or oFrmEvt.GrdSERVICOEVT.RowCount > 1 Then
         With oFrmEvt.GrdSERVICOEVT
            For i = 1 To .RowCount - 1
               If Val(.CellValue(i, "IDSERVICOEVT")) <> 0 And Val(.CellValue(i, "IDTPSERVICO")) <> 0 And Val(.CellValue(i, "IDTPTRATAMENTO")) <> 0 Then
                  If Not .ColVisible("IDAREA") Or Val(.CellValue(i, "IDAREA")) <> 0 Or Mid(.CellText(i, "IDAREA"), 1, 1) = "-" Then
                     TbServico.IDSERVICOEVT = Val(.CellValue(i, "IDSERVICOEVT"))
                     TbServico.IDTPSERVICO = Val(.CellValue(i, "IDTPSERVICO"))
                     TbServico.IDTPTRATAMENTO = Val(.CellValue(i, "IDTPTRATAMENTO"))
                     TbServico.IDAREA = Val(.CellValue(i, "IDAREA"))
                     
                     cQueries.Add TbServico.QrySave
                  End If
               End If
            Next
         End With
      End If
   End If
   
   If cQueries.Count > 0 Then
      MSalvarSessao = mvarSys.xDb.Executa(cQueries)
   End If
   Set cQueries = Nothing
   Set TbServico = Nothing
End Function
Private Function MCopiarSessao(ByRef pTbEvento As Object) As Boolean
   Dim i          As Long
   Dim cQueries   As Collection
   Dim TbServico  As Object 'TB_OSERVICOEVT
         
   MCopiarSessao = True
   
   If Not (oFrmEvt.ChkMeeting.Value = vbChecked) Then
      Call SetTag(oFrmEvt.GrdSERVICOEVT, "ISDIRT", "1")
   End If
   Set cQueries = New Collection
   
   'mvarIDCLIENTE =MSalvarCliente
   oFrmEvt.TxtIDCLIENTE = MSalvarCliente
   
   pTbEvento.IDCLIENTE = xVal(oFrmEvt.TxtIDCLIENTE.Text)
   pTbEvento.FLGCONFIRMADO = oFrmEvt.ChkFLGCONFIRMADO.Value
   pTbEvento.FLGCANCELADO = oFrmEvt.ChkFLGCANCELADO.Value
   pTbEvento.FLGREMARCADO = oFrmEvt.ChkFLGREMARCADO.Value
   'If pTbEvento.FLGCONFIRMADO = 1 Then
   '   pTbEvento.FLGCANCELADO = 0
   '   oFrmEvt.ChkFLGCANCELADO.Value
   'End If
   
   If pTbEvento.isDirt Then cQueries.Add pTbEvento.QrySave
   
   If GetTag(oFrmEvt.GrdSERVICOEVT, "ISDIRT", 0) = 1 Then
      
      Set TbServico = CriarObjeto("BANCO_3R.TB_OSERVICOEVT", False)
      Set TbServico.xDb = mvarSys.xDb
      TbServico.IDLOJA = pTbEvento.IDLOJA
      TbServico.IDEVENTO = pTbEvento.IDEVENTO

      cQueries.Add TbServico.QryDelete(Ch_IDLOJA:=TbServico.IDLOJA, Ch_IDEVENTO:=TbServico.IDEVENTO)
      If oFrmEvt.ChkMeeting.Value = vbChecked Then
         With oFrmEvt.GrdSERVICOEVT
            For i = 1 To .RowCount - 1
               If Val(.CellValue(i, "IDSERVICOEVT")) <> 0 And Val(.CellValue(i, "IDTPSERVICO")) <> 0 And Val(.CellValue(i, "IDTPTRATAMENTO")) <> 0 And Val(.CellValue(i, "IDAREA")) <> 0 Then
                  TbServico.IDSERVICOEVT = Val(.CellValue(i, "IDSERVICOEVT"))
                  TbServico.IDTPSERVICO = Val(.CellValue(i, "IDTPSERVICO"))
                  TbServico.IDTPTRATAMENTO = Val(.CellValue(i, "IDTPTRATAMENTO"))
                  TbServico.IDAREA = Val(.CellValue(i, "IDAREA"))
                  
                  cQueries.Add TbServico.QrySave
               End If
            Next
         End With
      End If
   End If
   If cQueries.Count > 0 Then
      MCopiarSessao = mvarSys.xDb.Executa(cQueries)
   End If
   Set cQueries = Nothing
   Set TbServico = Nothing
End Function
Public Function MProcuraCliente(pTEL As String, Optional pNome As String, Optional pIDCLIENTE As Long) As String
   Dim Sql        As String
   Dim MyRS       As Object
   Dim sTel       As String
   Dim bAchou     As Boolean
   Dim sNome      As String
   Dim nIDCLIENTE As Long
   Dim MyCliente  As Object
   Dim sTag       As String
   Dim sMsg      As String

   sTag = ""
'   Call SetTag(sTag, "IDCLIENTE", pIDCLIENTE)
'   Call SetTag(sTag, "NOME", pNome)
'   Call SetTag(sTag, "TEL1", pTEL)
'   Call SetTag(sTag, "TEL2", MyRs("TEL2") & "")
   sTel = Trim(Replace(pTEL, " ", ""))
   sTel = Trim(Replace(pTEL, ".", ""))
   sTel = "'" & Replace(pTEL, "-", "") & "'"
   sNome = Trim(pNome)
   If sTel = "''" Then Exit Function

   nIDCLIENTE = pIDCLIENTE
   pIDCLIENTE = 0

   Sql = "Select * " & vbNewLine
   Sql = Sql & " From OCLIENTE" & vbNewLine
   Sql = Sql & " Where (IDLOJA0=0 Or IDLOJA=" & mvarIDLOJA & ")" & vbNewLine
   Sql = Sql & " And (Replace(TEL1,'-','') =" & sTel & vbNewLine
   Sql = Sql & " Or Replace(TEL2,'-','') =" & sTel
   Sql = Sql & " )"
   
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      If MyRS.RecordCount = 1 Then
         If UCase(sNome) <> UCase(Trim(MyRS("NOME"))) And UCase(sNome) <> "" Then
            sMsg = "Existe um cliente com mesmo telefone." & vbNewLine
            sMsg = sMsg & vbNewLine
            sMsg = sMsg & "Nome : " & MyRS("NOME") & "" & vbNewLine
            If MyRS("TEL1") & "" <> "" Then
               sMsg = sMsg & "Tel.    : " & MyRS("TEL1") & "" & vbNewLine
            End If
            If MyRS("TEL2") & "" <> "" Then
               If MyRS("TEL1") & "" <> "" Then sMsg = sMsg & Space(5)
               sMsg = sMsg & "Cel. : " & MyRS("TEL2") & "" & vbNewLine
            End If
            If MyRS("EMAIL") & "" <> "" Then
               sMsg = sMsg & "e-Mail : " & MyRS("EMAIL") & "" & vbNewLine
            End If
            sMsg = sMsg & vbNewLine
            sMsg = sMsg & "Neste caso você só poderá alterar ou incluir um registro" & vbNewLine
            sMsg = sMsg & "na tela [Cadastro de Cliente]."
            
            Call ExibirAviso(sMsg)
         End If
         Call SetTag(sTag, "IDCLIENTE", MyRS("IDCLIENTE") & "")
         Call SetTag(sTag, "NOME", MyRS("NOME") & "")
         Call SetTag(sTag, "TEL1", MyRS("TEL1") & "")
         Call SetTag(sTag, "TEL2", MyRS("TEL2") & "")
         Call SetTag(sTag, "EMAIL", MyRS("EMAIL") & "")
         Call SetTag(sTag, "FOTOTIPO", MyRS("FOTOTIPO") & "")
         
         pIDCLIENTE = MyRS("IDCLIENTE")
         
      ElseIf MyRS.RecordCount > 1 Then
      
         If xVal(nIDCLIENTE) <> 0 Then
            '* Procurar ID na Lista.
            '* Se achar deve-se atualizar registro.
            '* Senão deve-se selecionar nome ou criar novo.
            'Set MyCliente = New TL_NovoLov
            Set MyCliente = CriarObjeto("Contato3R.TL_NovoLov", False)
            With MyCliente
               Set .Sys = mvarSys
               .Filtro = pTEL '"26139999"
               Call .ShowClientes
               If .IdSelecionado <> 0 Then
                  nIDCLIENTE = .IdSelecionado
               End If
            End With
            Set MyCliente = Nothing
            
            While Not MyRS.EOF
               If nIDCLIENTE = MyRS("IDCLIENTE") Then
               Call SetTag(sTag, "IDCLIENTE", MyRS("IDCLIENTE") & "")
               Call SetTag(sTag, "NOME", MyRS("NOME") & "")
               Call SetTag(sTag, "TEL1", MyRS("TEL1") & "")
               Call SetTag(sTag, "TEL2", MyRS("TEL2") & "")
               Call SetTag(sTag, "EMAIL", MyRS("EMAIL") & "")
               Call SetTag(sTag, "FOTOTIPO", MyRS("FOTOTIPO") & "")
                  
                  pIDCLIENTE = MyRS("IDCLIENTE")
                  bAchou = True
                  MyRS.MoveLast
               End If
               MyRS.MoveNext
            Wend
         End If
         If Not bAchou Then
            '* Deve-se selecionar nome ou criar novo
            '**** Obs. Se a opção for criar novo, 'pIDCLIENTE' deverá ser 0(zero)
            sMsg = "Vários clientes possuem este telefone." & vbNewLine
            sMsg = sMsg & vbNewLine
            sMsg = sMsg & "Favor informar outro telefone."
            Call ExibirStop(sMsg)
         End If
         
      End If
   Else
      pIDCLIENTE = nIDCLIENTE
   End If
'   If pIDCLIENTE = 0 And nIDCLIENTE = 0 Then
'      Call SetTag(sTAG, "IDCLIENTE", MyRs("IDCLIENTE") & "")
'      Call SetTag(sTAG, "NOME", MyRs("NOME") & "")
'      Call SetTag(sTAG, "TEL1", MyRs("TEL1") & "")
'      Call SetTag(sTAG, "TEL2", MyRs("TEL2") & "")
'      Call SetTag(sTAG, "EMAIL", MyRs("EMAIL") & "")
'      Call SetTag(sTAG, "FOTOTIPO", MyRs("FOTOTIPO") & "")
'   End If
   MProcuraCliente = sTag
End Function
Public Sub MReadRPatternExceptions(pPattern As CalendarRecurrencePattern)
   Dim Sql     As String
   Dim MyRS    As Object
   Dim pEvent  As CalendarEvent
    
   Sql = "Select * " & vbNewLine
   Sql = Sql & " From OEVENTOAGENDA" & vbNewLine
   Sql = Sql & " Where IDLOJA = " & mvarIDLOJA
   Sql = Sql & " And IDAGENDA= " & mvarIDAGENDA
   Sql = Sql & " And IDEVENTOREC = " & pPattern.Id
   Sql = Sql & " And RecurrenceState = " & xtpCalendarRecurrenceException
    
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      While Not MyRS.EOF
         Call MRecordsetParaEvento(MyRS, pEvent, True)
         If Not pEvent Is Nothing Then
            pPattern.SetException pEvent
         End If
         
         MyRS.MoveNext
      Wend
   End If
End Sub
Private Function GetSqlEventoDia(ByVal dtDay As Date, Optional bDia As Boolean) As String
   Dim nYear As Long
   Dim nMonth As Long
   Dim nDay As Long
   Dim Sql As String
   Dim i As Integer
    
   nYear = Year(dtDay)
   nMonth = Month(dtDay)
   nDay = Day(dtDay)
    
   Sql = "Select E.*, Day(E.StartDateTime) [DIA]" & vbNewLine
   Sql = Sql & ", A.IDATENDIMENTO, A.SITATEND, V.IDVENDA, V.SITVENDA" & vbNewLine
   Sql = Sql & " From OEVENTOAGENDA E" & vbNewLine
    'Sql = Sql & " And A.IDFUNCIONARIO=4" & vbNewLine
   Sql = Sql + " Left Join OATENDIMENTO A On E.IDLOJA=A.IDLOJA And E.IDEVENTO=A.IDEVENTO" & vbNewLine
   Sql = Sql + " Left Join OATENDIMENTO_VENDA AV On A.IDLOJA=AV.IDLOJA And A.IDATENDIMENTO=AV.IDATENDIMENTO" & vbNewLine
   Sql = Sql + " Left Join CVENDA V On AV.IDLOJA=V.IDLOJA And AV.IDVENDA=V.IDVENDA" & vbNewLine
    
   Sql = Sql & " Where E.IDLOJA=" & mvarIDLOJA & vbNewLine
   Sql = Sql & " And (E.RecurrenceState = " & xtpCalendarRecurrenceNotRecurring & vbNewLine
   Sql = Sql & "   Or E.RecurrenceState = " & xtpCalendarRecurrenceMaster & vbNewLine
   Sql = Sql & ")"
   Sql = Sql & " And Year(E.StartDateTime)  = " & nYear & vbNewLine
   Sql = Sql & " And Month(E.StartDateTime) = " & nMonth & vbNewLine
   If bDia Then Sql = Sql & " And Day(E.StartDateTime)   = " & nDay & vbNewLine
   If mvarFLGCANCELADO = 0 Then
      If mvarFLGAGENDADO = 1 Then
         Sql = Sql & " And FLGCANCELADO = 0" & vbNewLine
      End If
   ElseIf mvarFLGCANCELADO = 1 Then
      If mvarFLGAGENDADO = 0 Then
         Sql = Sql & " And FLGCANCELADO = 1" & vbNewLine
      End If
   End If
   'Sql = Sql & " And Year(E.EndDateTime)  >= " & nYear & vbnewline
   'Sql = Sql & " And Month(E.EndDateTime) >= " & nMonth & vbnewline
   'Sql = Sql & " And Day(E.EndDateTime)   >= " & nDay & vbnewline
   
   Sql = Sql & " And E.IDSALA in (" & mvarSys.Propriedades("SALAS") & ")" & vbNewLine
   
   Sql = Sql & " Order By E.StartDateTime"
    GetSqlEventoDia = Sql
End Function
'// The SQL script below can be created in 2 ways - simple or optimized.
'//
'// * The simple script is read all events from DB.
'//   RemindersManager will filter events and use only events with reminder
'//   which will be Fired untill dtFrom + PeriodMinutes.
'//   This can be very slow operation,
'//   but it is called no so often - once per 2 hours by default.
'//
'// * The optimized script is filter events instead of reminder manager.
'//   This can significantly improve performance.
'//
'// The simple variant can be useful for some debug purposes.
'//
'// NOTE: generally dtFrom is Now.
Public Function GetSqlEventoAlerta(ByVal dtFrom As Date, ByVal PeriodMinutes As Long)
   Dim Sql As String
   Dim dtUntil As Date
   Dim strUntil As String
   
   Sql = "Select E.* " & vbNewLine
   Sql = Sql & " From OEVENTOAGENDA E" & vbNewLine
   Sql = Sql & " Left Join OATENDIMENTO A On E.IDLOJA=A.IDLOJA And E.IDEVENTO=A.IDEVENTO"
   Sql = Sql & " Where A.IDEVENTO Is Null"
   Sql = Sql & " And E.FLGCANCELADO=0"
   Sql = Sql & " And E.IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And (E.RecurrenceState = " & xtpCalendarRecurrenceNotRecurring
   Sql = Sql & " OR E.RecurrenceState = " & xtpCalendarRecurrenceMaster & ")" & vbNewLine
   '------------------------------
   'If bOptimized Then
   '------------------------------
   If IsDate(dtFrom) Then
      dtUntil = DateAdd("n", PeriodMinutes, dtFrom)
   End If
   strUntil = mvarSys.xDb.SqlDate(dtUntil, 3)
          
   Sql = Sql & " AND (E.IsReminder <> 0 " & vbNewLine
   Sql = Sql & "     OR ( E.RecurrenceState = " & xtpCalendarRecurrenceMaster & vbNewLine
   Sql = Sql & "          AND E.IDEVENTOREC IN (Select E2.IDEVENTOREC" & vbNewLine
   Sql = Sql & "                                      From OEVENTOAGENDA E2" & vbNewLine
   Sql = Sql & "                                      Where E2.IDLOJA= " & mvarIDLOJA
   Sql = Sql & "                                      And E2.RecurrenceState = " & xtpCalendarRecurrenceException & vbNewLine
   Sql = Sql & "                                      And E2.IsReminder <> 0 " & vbNewLine
   Sql = Sql & "                                      And DATEADD(""n"", -1 * E2.ReminderMinutesBeforeStart, E2.StartDateTime) <= " & strUntil & vbNewLine
   Sql = Sql & "                                      )" & vbNewLine
   Sql = Sql & "         ) " & vbNewLine
   Sql = Sql & "      ) " & vbNewLine
   Sql = Sql & " AND DATEADD(""n"", -1 * E.ReminderMinutesBeforeStart, E.StartDateTime) <= " & strUntil
   '------------------------------
   'End If
   '------------------------------
   
   GetSqlEventoAlerta = Sql
End Function

Private Sub mvarCalendar_PrePopulate(ByVal ViewGroup As XtremeCalendarControl.CalendarViewGroup, ByVal Events As XtremeCalendarControl.CalendarEvents)
   Dim pEvent  As CalendarEvent
   For Each pEvent In Events
      Call SetCustomIcons(pEvent)
   Next
   'mvarCalendar.RedrawControl
End Sub
Public Sub GSetCustomEvent(ByRef pEvent As CalendarEvent, mvarSys As Object, _
                                    Optional pFLGCONFIRMADO, _
                                    Optional pFLGCANCELADO, _
                                    Optional pFLGREMARCADO, _
                                    Optional pIDATENDIMENTO, _
                                    Optional pSITATEND, _
                                    Optional pIDVENDA, _
                                    Optional pSITVENDA, _
                                    Optional pIDCLIENTE, _
                                    Optional pIDSALA)
   Call SetCustomEvent(pEvent, mvarSys, pFLGCONFIRMADO, pFLGCANCELADO, pFLGREMARCADO, pIDATENDIMENTO, pSITATEND, pIDVENDA, pSITVENDA, pIDCLIENTE, pIDSALA)
End Sub
Public Sub SetEventFlag(ByRef pEvent As CalendarEvent, Optional pRs As Object)   'Optional pIDEVENTO, Optional pIDATEND)
   Call SetEventFlag2(pEvent, pRs)
End Sub
Public Sub SetCustomIcons(ByRef pEvent As XtremeCalendarControl.CalendarEvent)
   Call SetCustomIcons2(pEvent)
End Sub
Private Function VerificarRemarcacao(pIDLOJA As Integer, pIDEVENTO As Integer) As Boolean
   Dim TbEvt   As Object
   Dim TbEvt2  As Object
   Dim Sql     As String
   
   Set TbEvt = CriarObjeto("BANCO_3R.TB_OEVENTOAGENDA")
   With TbEvt
      Set .xDb = mvarSys.xDb
      If .Pesquisar(Ch_IDLOJA:=pIDLOJA, Ch_IDEVENTO:=pIDEVENTO) Then
         Set TbEvt2 = CriarObjeto("BANCO_3R.TB_OEVENTOAGENDA")
         TbEvt2.xDb = mvarSys.xDb
         Sql = ""
         Sql = Sql & " IDLOJA=" & .IDLOJA
         Sql = Sql & " And IDCLIENTE=" & .IDCLIENTE
         Sql = Sql & " And FLGCANCELADO=0"
         Sql = Sql & " And StartDateTime>=" & SqlDate(.StartDateTime)
         
         VerificarRemarcacao = TbEvt2.Pesquisar(Ch_Where:=Sql)
         Set TbEvt2 = Nothing
      End If
      Set TbEvt = Nothing
   End With
   
   

End Function
Private Sub mvarCalendar_PrePopulateDay(ByVal ViewDay As XtremeCalendarControl.CalendarViewDay)
   'Call SetCustomEvent(mvarCalEvent, mvarSys, pFLGCONFIRMADO:=TbEvt.FLGCONFIRMADO, pFLGCANCELADO:=TbEvt.FLGCANCELADO, pIDATENDIMENTO:=mvarIDATENDIMENTO, pIDCLIENTE:=mvarIDCLIENTE)
   'Call SetCustomIcons(pEvent)
'    If me.CalendarControl.MultipleResources.Count > 1 Then
'
'        Dim arRes As New CalendarResources
'        'arRes.RemoveAll
            
'        If Day(ViewDay.Date) Mod 2 = 0 Then
'            arRes.Add me.CalendarControl.MultipleResources.Item(0)
'            ViewDay.SetMultipleResources arRes
'        ElseIf Day(ViewDay.Date) Mod 3 = 0 Then
'            arRes.Add me.CalendarControl.MultipleResources.Item(me.CalendarControl.MultipleResources.Count - 1)
'            ViewDay.SetMultipleResources arRes
'        End If
'    End If
End Sub

Private Sub oFrmEvt_ChkMeetingClick()
   'Dim TbAgen  As Object 'TB_OAGENDA
   Dim TbSala  As Object 'TB_OSALA
   With oFrmEvt
      If Not .TxtLocation.Enabled Then
         .TxtLocation.Text = .CmbSchedule.Text
         'Set TbSala = CriarObjeto("BANCO_3R.TB_OSALA")
         'Set TbSala.xDb = mvarSys.xDb
         ''Set TbAgen = CriarObjeto("BANCO_3R.TB_OAGENDA")
         ''Set TbAgen.xDb = mvarSys.xDb
         ''If TbAgen.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDAGENDA:=mvarIDAGENDA) Then
         '   If TbSala.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDSALA:=1) Then ' TbAgen.IDSALA) Then
         '      .TxtLocation.Text = "Sala " & TbSala.CODSALA
         '   End If
         ''End If
         'Set TbSala = Nothing
         ''Set TbAgen = Nothing
      End If
   End With
End Sub

Private Sub oFrmEvt_CmbNOMEClick()
   Call PopulaCliente(oFrmEvt.CmbNOME.ItemData(oFrmEvt.CmbNOME.ListIndex))
End Sub

Private Sub oFrmEvt_CmdIDCLIENTEClick()
   Call ShowCliente(mvarSys, mvarIDLOJA, xVal(oFrmEvt.TxtIDCLIENTE.Text))
End Sub
Private Sub oFrmEvt_CmdLovClick()
   Screen.MousePointer = vbHourglass
   Call PopulaCliente
   Screen.MousePointer = vbDefault
End Sub
Public Sub PopulaCliente(Optional pIDCLIENTE As Long)
   Dim MyCliente As Object

   On Error GoTo Saida
   
   If pIDCLIENTE = 0 Then
      Set MyCliente = CriarObjeto("Contato3R.TL_NovoLov", False)
      With MyCliente
         Set .Sys = mvarSys
         .Filtro = Trim(Trim(oFrmEvt.TxtNOME.Text))
         If Trim(Trim(oFrmEvt.TxtNOME.Text)) = "" Then
            .Filtro = Trim(Trim(oFrmEvt.TxtNOME.Text) + " " + Trim(oFrmEvt.TxtTEL1.Text) + " " + Trim(oFrmEvt.TxtTEL2.Text)) '+ " " + Trim(oFrmEvt.TxtEMAIL.Text))
         End If
         Screen.MousePointer = vbDefault
         Call .ShowClientes
         If .IdSelecionado <> 0 Then
            pIDCLIENTE = .IdSelecionado
         End If
      End With
   End If
   Screen.MousePointer = vbHourglass
   If pIDCLIENTE > 0 Then
      Dim MyRS As Object
      Dim Sql As String
      Sql = Sql_OCLIENTE(mvarIDLOJA, pIDCLIENTE)
      If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
         oFrmEvt.TxtIDCLIENTE.Text = pIDCLIENTE
         oFrmEvt.TxtIDCLIENTE.Text = MyRS("IDCLIENTE")
         oFrmEvt.TxtNOME.Text = MyRS("NOME")
         oFrmEvt.TxtTEL1.Text = MyRS("TEL1")
         oFrmEvt.TxtTEL2.Text = MyRS("TEL2")
         oFrmEvt.TxtEMAIL.Text = MyRS("EMAIL")
         If MyRS("OBS") <> "" Then
            oFrmEvt.CmdIDCLIENTE.Caption = "Cliente *"
         End If
         Call CarregarSessoes
      End If
      oFrmEvt.GrdSERVICOEVT.SetFocus
   ElseIf pIDCLIENTE < 0 Then
      Screen.MousePointer = vbDefault
      Call ExibirAviso(" Ocorreu falha na pesquisa.")
   End If

   Set MyCliente = Nothing
   Exit Sub
Saida:
   If Err > 0 Then
      Screen.MousePointer = vbDefault
      MsgBox Err & " - " & Error
   End If
End Sub
Private Sub CarregarSessoes()
   'Dim MySessao   As Object
   'Dim MyEvento   As Object
   'Dim MyArea     As Object
   'Dim MyCli      As Object
   
   Dim Sql        As String
   Dim nTempo     As Integer
   Dim nResto     As Integer
   Dim pLabel     As CalendarEventLabel
   Dim sTag       As String
   Dim i          As Integer
   Dim nID        As Long
   Dim sSexo      As String
   Dim nIDTPSERVICO As Integer
   Dim nIDMANIPULO As Integer
   
   Dim MyRsAtend  As Object
   Dim MyRsSessao As Object
   Dim MyRsEvento As Object
   Dim MyRsArea   As Object
   Dim MyRsTPMAN  As Object
   Dim MyRsCli    As Object
      
   On Error GoTo Saida
   
   'If mvarIDCLIENTE = 0 Then Exit Sub
   If oFrmEvt.GrdSERVICOEVT.RowCount > 2 Then
      For i = oFrmEvt.GrdSERVICOEVT.RowCount - 1 To i Step -1
         If xVal(oFrmEvt.GrdSERVICOEVT.CellValue(i, "IDTPSERVICO")) = 0 Or xVal(oFrmEvt.GrdSERVICOEVT.CellValue(i, "IDTPTRATAMENTO")) = 0 Then
            oFrmEvt.GrdSERVICOEVT.RemoveRow i
         End If
      Next
   End If

   If oFrmEvt.GrdSERVICOEVT.RowCount > 2 Then
      Exit Sub
   ElseIf oFrmEvt.GrdSERVICOEVT.RowCount = 2 Then
      If xVal(oFrmEvt.GrdSERVICOEVT.CellValue(1, "IDTPSERVICO")) <> 0 And xVal(oFrmEvt.GrdSERVICOEVT.CellValue(1, "IDTPTRATAMENTO")) <> 0 Then
         Exit Sub
      End If
   End If
   
   Call SetTag(oFrmEvt.GrdSERVICOEVT, "CARREGANDO", 1)
   Sql = Sql_OATENDIMENTO(mvarIDLOJA, pIDCLIENTE:=xVal(oFrmEvt.TxtIDCLIENTE.Text))
   Sql = Sql & " Order By DTATEND desc"
   If mvarSys.xDb.AbreTabela(Sql, MyRsAtend) Then
      Sql = Sql_OCLIENTE(mvarIDLOJA, xVal(oFrmEvt.TxtIDCLIENTE.Text))
      sSexo = "F"
      If mvarSys.xDb.AbreTabela(Sql, MyRsCli) Then
         sSexo = IIf(Trim(MyRsCli("SEXO")) = "", "F", MyRsCli("SEXO"))
      End If
      If (MyRsAtend("HHINI") & "") = "" Or (MyRsAtend("HHFIM") & "") = "" Then
         nTempo = 0
      Else
         nTempo = DateDiff("n", (MyRsAtend("HHINI") & ""), (MyRsAtend("HHFIM") & ""))
      End If
      
      Sql = Sql_OSESSAO(mvarIDLOJA, pIDATENDIMENTO:=MyRsAtend("IDATENDIMENTO"))
      If mvarSys.xDb.AbreTabela(Sql, MyRsSessao) Then
         nIDTPSERVICO = MyRsSessao("IDTPSERVICO").Value
         If nIDTPSERVICO = 1 Then
         
         Else
            nResto = 5 - nTempo Mod 5
            nTempo = nTempo + IIf(nResto < 5, nResto, 0)
            If IsDate(oFrmEvt.CmbStartTime.Text) Then
               oFrmEvt.CmbEndTime.Text = DateAdd("n", nTempo, oFrmEvt.CmbStartTime.Text)
            End If
         End If
         
         nTempo = 0
         Call SetTag(oFrmEvt.GrdSERVICOEVT, "ISDIRT", 1)
         oFrmEvt.GrdSERVICOEVT.Clear
         While Not MyRsSessao.EOF
            nIDTPSERVICO = MyRsSessao("IDTPSERVICO").Value
            If nIDTPSERVICO = 1 Then
               nIDTPSERVICO = 2
               nIDMANIPULO = xVal(MyRsSessao("IDMANIPULO") & "")
               If nIDMANIPULO = 0 Then nIDMANIPULO = 1
               Sql = Sql_OMANIPULO(mvarIDLOJA, nIDMANIPULO)
               If mvarSys.xDb.AbreTabela(Sql, MyRsTPMAN) Then
                  Sql = Sql_OAREA_TRATAMENTO(mvarIDLOJA, MyRsSessao("IDAREA"), pIDTPTRATAMENTO:=MyRsSessao("IDTPTRATAMENTO"), pIDTPMANIPULO:=nIDMANIPULO)
                  If mvarSys.xDb.AbreTabela(Sql, MyRsArea) Then
                     nTempo = nTempo + (MyRsArea("TEMPOMED") * IIf(sSexo = "F", 1, MyRsArea("INDICEH")))
                  End If
               End If
            End If
            If nIDTPSERVICO > 1 Then
               oFrmEvt.ChkFLGAVALIACAO.Value = vbUnchecked
            End If
            mvarTlEvt.IncluiGrdLinha
            If oFrmEvt.GrdSERVICOEVT.RowCount = 1 Then
               mvarTlEvt.IncluiGrdLinha
            End If
            With oFrmEvt.GrdSERVICOEVT
               i = .RowCount - 1
               If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
               Call SetTag(sTag, "IDSERVICOEVT", i)
               Call SetTag(sTag, "IDTPSERVICO", nIDTPSERVICO)
               Call SetTag(sTag, "IDTPTRATAMENTO", MyRsSessao("IDTPTRATAMENTO").Value)
               Call SetTag(sTag, "IDAREA", MyRsSessao("IDAREA").Value)
               .RowTag(i) = sTag
               
               .CellValue(i, "IDSERVICOEVT") = i
               .CellValue(i, "IDTPSERVICO") = nIDTPSERVICO
               .CellValue(i, "IDTPTRATAMENTO") = MyRsSessao("IDTPTRATAMENTO").Value
               .CellValue(i, "IDAREA") = MyRsSessao("IDAREA").Value
               
               If i = 1 Then
                  oFrmEvt.CmbShowTimeAs.Enabled = False
                  oFrmEvt.CmbShowTimeAs.ListIndex = DefineShowAs(nIDTPSERVICO)
                  
                  nID = DefineLabelID(MyRsSessao("IDTPTRATAMENTO"), nIDTPSERVICO)
                  Set pLabel = mvarCalendar.DataProvider.LabelList.Find(nID)
                  If Not pLabel Is Nothing Then
                     oFrmEvt.ctrlColor.BackColor = pLabel.Color
                  End If
               End If
               
            End With
            MyRsSessao.MoveNext
         Wend
         If nTempo > 0 Then
            nResto = 5 - nTempo Mod 5
            nTempo = nTempo + IIf(nResto < 5, nResto, 0)
            If IsDate(oFrmEvt.CmbStartTime.Text) Then
               oFrmEvt.CmbEndTime.Text = DateAdd("n", nTempo, oFrmEvt.CmbStartTime.Text)
            End If
         End If
      End If
   Else
      Sql = Sql_OEVENTOAGENDA(mvarIDLOJA)
      Sql = Sql & " IDCLIENTE=" & xVal(oFrmEvt.TxtIDCLIENTE.Text)
      Sql = Sql & " Order By StartDateTime desc "
      If mvarSys.xDb.AbreTabela(Sql, MyRsEvento) Then
         nTempo = DateDiff("n", MyRsEvento("StartDateTime"), MyRsEvento("EndDateTime"))
         nResto = 5 - nTempo Mod 5
         nTempo = nTempo + IIf(nResto < 5, nResto, 0)
         If IsDate(oFrmEvt.CmbStartTime.Text) Then
            oFrmEvt.CmbEndTime.Text = DateAdd("n", nTempo, oFrmEvt.CmbStartTime.Text)
         End If
      
         Call SetTag(oFrmEvt.GrdSERVICOEVT, "ISDIRT", 1)
         Call MPopulaServicoEVT(MyRsEvento("IDEVENTO"), False)
      End If
   End If
   Call SetTag(oFrmEvt.GrdSERVICOEVT, "CARREGANDO", 0)
   Exit Sub
Saida:
   If Err > 0 Then
      Screen.MousePointer = vbDefault
      MsgBox Err & " - " & Error
   End If
End Sub

Private Sub oFrmEvt_GrdTarefasDblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   'If (lRow = oFrmEvt.GrdTarefas.RowCount) Then
   '   bRequestEdit = False
   'Else
      If lRow > 0 Then
         Call EditarTarefa(xVal(oFrmEvt.GrdTarefas.CellValue(lRow, "IDTAREFA")))
         Call MPopulaTarefas(mvarTlEvt.IDEVENTO)
      End If
   'End If
End Sub
Private Sub EditarTarefa(pIDTAREFA As Long)
   Dim TlCadastro As Object
      
   If pIDTAREFA > 0 Then
      Set TlCadastro = CriarObjeto("Tarefa3R.TL_CadTarefa")
      With TlCadastro
         Set .Sys = mvarSys
         .IDLOJA = mvarIDLOJA
         .IDTAREFA = pIDTAREFA
         Call .Show
      End With
   End If
End Sub
Private Sub oFrmEvt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF1 Then
      If oFrmEvt.ActiveControl Is oFrmEvt.TxtNOME Then
         Call oFrmEvt_CmdLovClick
      End If
   ElseIf KeyCode = vbKeyReturn Then
      If oFrmEvt.ChkMeeting.Value = vbChecked Then
         If xVal(oFrmEvt.TxtIDCLIENTE.Text) = 0 Then
            Call oFrmEvt_CmdLovClick
         End If
      End If
   End If
End Sub
Private Sub oFrmEvt_TxtTEL1LostFocus()
   Dim sTag As String
   sTag = MProcuraCliente(oFrmEvt.TxtTEL1.Text, oFrmEvt.TxtNOME.Text, xVal(oFrmEvt.TxtIDCLIENTE.Text))
   If sTag <> "" Then
      'mvarIDCLIENTE = GetTag(sTag, "IDCLIENTE", "0")
      oFrmEvt.TxtIDCLIENTE.Text = GetTag(sTag, "IDCLIENTE", "0")
      oFrmEvt.TxtNOME.Text = GetTag(sTag, "NOME", "")
      oFrmEvt.TxtTEL1.Text = GetTag(sTag, "TEL1", "")
      oFrmEvt.TxtTEL2.Text = GetTag(sTag, "TEL2", "")
      oFrmEvt.TxtEMAIL.Text = GetTag(sTag, "EMAIL", "")
      oFrmEvt.GrdSERVICOEVT.SetFocus
   End If
End Sub
Private Sub oFrmEvt_TxtTEL2LostFocus()
   Dim sTag As String
   sTag = MProcuraCliente(oFrmEvt.TxtTEL2.Text, oFrmEvt.TxtNOME.Text, xVal(oFrmEvt.TxtIDCLIENTE.Text))
   If sTag <> "" Then
      'mvarIDCLIENTE = GetTag(sTag, "IDCLIENTE", "0")
      oFrmEvt.TxtIDCLIENTE.Text = GetTag(sTag, "IDCLIENTE", "0")
      oFrmEvt.TxtNOME.Text = GetTag(sTag, "NOME", "")
      oFrmEvt.TxtTEL1.Text = GetTag(sTag, "TEL1", "")
      oFrmEvt.TxtTEL2.Text = GetTag(sTag, "TEL2", "")
      oFrmEvt.TxtEMAIL.Text = GetTag(sTag, "EMAIL", "")
      oFrmEvt.GrdSERVICOEVT.SetFocus
   End If
End Sub
Public Function VerificaEvento(ByRef pEvent As XtremeCalendarControl.CalendarEvent, bDelete As Boolean, Optional bExibeMsg As Boolean = True) As Boolean
   Dim Sql  As String
   Dim MyRS As Object
   Dim bAtendFechado As Boolean
   Dim bVendaFechada As Boolean
   Dim bExisteAtend  As Boolean
   Dim bExisteVenda  As Boolean
   Dim sMsg          As String
   Dim pIDEVENTO     As Integer
   
   If pEvent Is Nothing Then Exit Function
   pIDEVENTO = pEvent.Id

'   If xVal(pEvent.CustomProperties("FLGCANCELADO")) = 1 Then Exit Function
   
   If Not pEvent.CustomProperties("EXIBEMSG") Then
      If Not pEvent.CustomProperties("EXIBEMSG") = Empty Then
         bExibeMsg = False
      End If
   End If
     
   Sql = "Select A.IDATENDIMENTO, A.SITATEND, V.IDVENDA, V.SITVENDA" & vbNewLine
   Sql = Sql + " From OATENDIMENTO A" & vbNewLine
   Sql = Sql + " Left Join OATENDIMENTO_VENDA AV On A.IDLOJA=AV.IDLOJA And A.IDATENDIMENTO=AV.IDATENDIMENTO" & vbNewLine
   Sql = Sql + " Left Join CVENDA V On AV.IDLOJA=V.IDLOJA And AV.IDVENDA=V.IDVENDA" & vbNewLine
   Sql = Sql + " Where A.IDEVENTO=" & xVal(pIDEVENTO) & vbNewLine
   Sql = Sql + " And A.IDLOJA=" & xVal(mvarIDLOJA) & vbNewLine
   
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      bExisteAtend = True
      bAtendFechado = (MyRS("SITATEND") & "" = gSITATEND_Fechado)
      bExisteVenda = (xVal(MyRS("IDVENDA") & "") <> 0)
      If pEvent.CustomProperties("VENDAFECHADA") <> "" Then
         bVendaFechada = pEvent.CustomProperties("VENDAFECHADA")
      Else
         bVendaFechada = (MyRS("SITVENDA") & "" = gSITVENDA_Fechada)
      End If
   Else
      bExisteAtend = False
      bAtendFechado = False
      bVendaFechada = False
      bExisteVenda = False
   End If
   
   If bAtendFechado Then
      If bExibeMsg Then ExibirAviso ("Evento possui atendimento fechado e não pode ser alterado.")
      GoTo Saida
   End If
   If bExisteVenda Then
      If bVendaFechada Then
         If bExibeMsg Then ExibirAviso ("Evento possui venda fechada e não pode ser alterado.")
         GoTo Saida
      ElseIf bDelete Then
         Call ExibirAviso("Evento possui Venda e não pode ser alterado.")
         GoTo Saida
      End If
   End If
   
   If bDelete Then
      If bExisteAtend Then
         If bExibeMsg Then ExibirAviso ("Evento possui atendimento associado e não pode ser excluído.")
        GoTo Saida
      End If
   End If
   
   If pEvent.MeetingFlag = True And Not bDelete And pEvent.CustomProperties("FLGCANCELADO") <> "1" Then
      If VerificaSobreposicao(pEvent.Id, xVal(pEvent.CustomProperties("IDSALA")), pEvent.StartTime, pEvent.EndTime, bExibeMsg) Then
         GoTo Saida
      End If
   End If
   VerificaEvento = True
   pEvent.CustomProperties("EXIBEMSG") = True
   Exit Function
Saida:
   VerificaEvento = False
   pEvent.CustomProperties("EXIBEMSG") = True
End Function
Public Function VerificaSobreposicao(nIDEVENTO As Long, pIDSALA As Integer, sDateIni As String, sDateFim As String, Optional bExibeMsg As Boolean = True) As Boolean
   Dim Sql  As String
   Dim MyRS As Object
   Dim sMsg As String
   Dim nTSobrePos As Integer
      
   DoEvents
  
   If Format(mvarSys.xDb.Sysdate(), "dd/mm/yy") >= Format(sDateIni, "dd/mm/yy") Then Exit Function
   
      
   sDateIni = Format(sDateIni, "dd/mm/yyyy hh:mm")
   sDateFim = Format(sDateFim, "dd/mm/yyyy hh:mm")
      
   Sql = ""
   Sql = Sql & "Select IDEVENTO, SUBJECT, STARTDATETIME, ENDDATETIME "
   Sql = Sql & " From OEVENTOAGENDA"
   Sql = Sql & " Where ISMEETING = 1"
   Sql = Sql & " And FLGCANCELADO<>1"
   Sql = Sql & " And IDLOJA=" & SqlNum(mvarIDLOJA)
   Sql = Sql & " And IDSALA =" & SqlNum(pIDSALA)
   Sql = Sql & " And IDEVENTO<>" & SqlNum(nIDEVENTO)
   Sql = Sql & " And ("
   Sql = Sql & " (StartDateTime <=" & SqlDate(sDateIni) & " And EndDateTime >" & SqlDate(sDateIni) & ")"
   Sql = Sql & " Or (StartDateTime <" & SqlDate(sDateFim) & " And EndDateTime >=" & SqlDate(sDateFim) & ")"
   Sql = Sql & " Or (StartDateTime >=" & SqlDate(sDateIni) & " And EndDateTime <=" & SqlDate(sDateFim) & ")"
   Sql = Sql & "     )"
   Sql = Sql & " Order by StartDateTime"
   sMsg = ""
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      MyRS.MoveFirst
      sMsg = vbNewLine
      While Not MyRS.EOF
         sMsg = sMsg & vbNewLine
         sMsg = sMsg & MyRS("IDEVENTO") & " - " & MyRS("SUBJECT") & vbNewLine
         sMsg = sMsg & " (" & Format(MyRS("STARTDATETIME"), "dd/mm/yyyy")
         sMsg = sMsg & " " & Format(MyRS("STARTDATETIME"), "hh:mm")
         sMsg = sMsg & " - " & Format(MyRS("ENDDATETIME"), "hh:mm")
         sMsg = sMsg & ")"
         sMsg = sMsg & vbNewLine
         MyRS.MoveNext
      Wend
      If MyRS.RecordCount = 1 Then
         '***********
         '* Tempo que premite sobreposição em minutos
         nTSobrePos = 5
         
         MyRS.MoveFirst
         If DateDiff("n", CDate(MyRS("STARTDATETIME")), CDate(sDateFim)) <= nTSobrePos Then
            VerificaSobreposicao = False
            Exit Function
         End If
         If DateDiff("n", CDate(sDateIni), CDate(MyRS("ENDDATETIME"))) <= nTSobrePos Then
            VerificaSobreposicao = False
            Exit Function
         End If
         If CDate(Format(sDateFim, "hh:mm")) < CDate(Format(MyRS("ENDDATETIME"), "hh:mm")) Then
            'sMsg = "O Atendimento irá sobrepor o horário do evento abaixo." & vbNewLine & vbNewLine & sMsg
            'sMsg = sMsg & "Tente marcar até xx?" & sMsg
            
            sMsg = "O Atendimento irá sobrepor o horário do evento abaixo." & vbNewLine & vbNewLine & sMsg
            Call ExibirInformacao(sMsg)
            'sMsg = sMsg & "Deseja atrasar o próximo evento?" & sMsg
            'If ExibirPergunta(sMsg) = vbYes Then
            '   Sql = "Update OEVENTOAGENDA"
            '   Sql = Sql & " Set StartDateTime=" & SqlDate(sDateFim)
            '   Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
            '   Sql = Sql & " And IDEVENTO=" & SqlNum(MyRs("IDEVENTO"))
            '   If mvarsys.xDb.Executa(Sql) Then
            '      Sql = "Update OATENDIMENTO"
            '      Sql = Sql & " Set HHINI=" & SqlDate(sDateFim)
            '      Sql = Sql & " , DTATEND=" & SqlDate(sDateFim)
            '      Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
            '      Sql = Sql & " And IDEVENTO=" & SqlNum(MyRs("IDEVENTO"))
            '      Call mvarsys.xDb.Executa(Sql)
            '   End If
               VerificaSobreposicao = True
               Exit Function
            'End If
         Else
            If bExibeMsg Then ExibirAviso ("O Atendimento irá sobrepor o horário do evento abaixo." & sMsg)
         End If
      Else
         If bExibeMsg Then ExibirAviso ("O Atendimento irá sobrepor os horários dos eventos abaixo." & sMsg)
      End If
      VerificaSobreposicao = True
   Else
      VerificaSobreposicao = False
   End If
End Function
Public Sub ExibirAtendimento(pCod As Long, Optional pCalControl)
   Dim TBAtend As Object
   Dim TlAtend As Object
   
   Set TBAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
   Set TBAtend.xDb = mvarSys.xDb
   If TBAtend.Pesquisar(Ch_IDLOJA:=mvarSys.Propriedades("IDLOJA"), Ch_IDATENDIMENTO:=pCod) Then
      Set TlAtend = CriarObjeto("Calendario3R.TL_Atendimento")
      With TlAtend
         Set .Sys = mvarSys
         If Not IsMissing(pCalControl) Then
            If Not pCalControl Is Nothing Then
               Set .CalControl = pCalControl
               Set .CalEvent = pCalControl.DataProvider.GetEvent(TBAtend.IDEVENTO)
            End If
         End If
         .IDEVENTO = TBAtend.IDEVENTO
         .IDATENDIMENTO = pCod
         .Show vbModal
      End With
   End If
   
   Set TBAtend = Nothing
   Set TlAtend = Nothing
End Sub
Public Sub ExibirEvento(pIDEVENTO As Long, Optional pCalControl)
   Dim TLEvento As TL_Evento
   Dim TlAtend As TL_Atendimento
   
   If ExisteAtendimento(pIDEVENTO) <> 0 Then
      Set TlAtend = New TL_Atendimento
      With TlAtend
         Set .Sys = mvarSys
         If Not IsMissing(pCalControl) Then
            Set .CalControl = pCalControl
            Set .CalEvent = pCalControl.DataProvider.GetEvent(pIDEVENTO)
         End If
         .IDEVENTO = pIDEVENTO
         .Show vbModal
      End With
   
   Else
      Set TLEvento = New TL_Evento
      Set mvarTlEvt = TLEvento
      If IsMissing(pCalControl) Then
         TLEvento.IDEVENTO = pIDEVENTO
      Else
         Set TLEvento.CalControl = pCalControl
         Set TLEvento.CalEvent = pCalControl.DataProvider.GetEvent(pIDEVENTO)
      End If
      Set TLEvento.Sys = mvarSys
      Set oFrmEvt = TLEvento.Form
      TLEvento.Show
   End If
End Sub
Public Sub SaveTheme()
    Dim px As PropExchange
    On Error Resume Next
    Set px = XtremeCalendarControl.CreatePropExchange
    
    If px.CreateAsXML(False, "CalendarThemeOffice2007") Then
        mvarCalendar.Theme.DoPropExchange px
        px.SaveToFile mvarSys.ExePath & "\cfgCalendarThemeOffice2007.xml"
    End If
End Sub




