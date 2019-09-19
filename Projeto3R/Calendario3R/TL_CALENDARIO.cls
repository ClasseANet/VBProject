VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CALENDARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim gDebug As Boolean

Dim NG_Cal As NG_Calendario

Private mvarToolBar   As CommandBar
Private WithEvents mvarCmmdBars  As XtremeCommandBars.CommandBars
Attribute mvarCmmdBars.VB_VarHelpID = -1
Private WithEvents mvarMe        As FrmCalendario
Attribute mvarMe.VB_VarHelpID = -1
Private WithEvents mvarPane      As FrmPaneCalendario
Attribute mvarPane.VB_VarHelpID = -1

Private Enum eMenuCal
   Mnu_Calendario = 2100
   Mnu_CalLembrete
   Mnu_CalOpcao
   Mnu_CalOpcaoAvancada
   
   'Mnu_CalEdit
   Mnu_CalEditOpen
   Mnu_CalEditDelete
   Mnu_CalEditDelete2
   Mnu_CalEditCliente
   Mnu_CalEditVenda
   Mnu_CalEditConfirm
   Mnu_CalEditCancel
   'Mnu_CalNew
   Mnu_CalNewEvent
   'Mnu_CalTime
   Mnu_CalTimeChange
   Mnu_CalTime60
   Mnu_CalTime30
   Mnu_CalTime15
   Mnu_CalTime10
   Mnu_CalTime5
End Enum
Private Enum eBarCal
   Bar_CalDia = 200
   Bar_CalDiaUtil
   Bar_CalSemana
   Bar_CalMes
   Bar_CalPgSetup
   Bar_CalPrintPreview
   Bar_CalPrint
End Enum
Private Enum ePopCal
   Pop_CalEdit = 100
   Pop_CalNew
   Pop_CalTime
End Enum

Dim Fx1 As Integer
Dim Fx2 As Integer
Dim Fx3 As Integer
Dim nTot As Integer

Dim gUnidades As Collection

Private mvarSys      As Object
Private mvarIDLOJA   As Integer
Private mvarTLEvento As TL_Evento
Private mvarTLAtend  As TL_Atendimento
Public Property Get NgCal() As NG_Calendario
   Set NgCal = NG_Cal
End Property
Public Property Get TLEvento() As TL_Evento
   If mvarTLEvento.CalControl Is Nothing Then
      Set mvarTLEvento.CalControl = mvarMe.CalendarControl
   End If
   Set TLEvento = mvarTLEvento
End Property
Public Property Get FrmCal() As Object
   Set FrmCal = mvarMe
End Property

Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   Set gSys = mvarSys
   
   mvarIDLOJA = xVal(mvarSys.Propriedades("IDLOJA"))
   If Not NgCal Is Nothing Then
      Set NgCal.Sys = mvarSys
   End If
   If Not mvarTLEvento Is Nothing Then Set mvarTLEvento.Sys = mvarSys
   If Not mvarTLAtend Is Nothing Then Set mvarTLAtend.Sys = mvarSys
   
   Set mvarCmmdBars = mvarSys.MDI.CommandBars
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Property Get Pane() As Object
    Set Pane = mvarPane
End Property
Public Sub Show(Optional Reload As Boolean = False)
   If gDebug Then MsgBox "TL_CALENDARIO.Show"
   
   If Reload Then
      Set Sys = Sys
      'Unload mvarMe
   End If

   Load mvarMe
   Call AcoplarForm(mvarMe, 2, mvarSys)
   mvarMe.Show
   
   '* Acoplar Form no MDI
   'If mvarSys.MDI.DockingPaneManager.Panes(2).Handle <> mvarMe.hwnd Then
   '   SetMDI mvarMe.hwnd, mvarSys.MDI.hwnd
   '   mvarSys.MDI.DockingPaneManager.Panes(2).Handle = mvarMe.hwnd
   'End If

   '* Exibir Menu no MDI
   If Not mvarSys.MDI.CommandBars.FindControl(, eMenuCal.Mnu_Calendario, False) Is Nothing Then
      mvarSys.MDI.CommandBars.FindControl(, eMenuCal.Mnu_Calendario).Visible = True
   End If
   
   If xVal(mvarSys.Propriedades("PONTO")) = 0 Then
      Call DockBarRightOf("BAR_CALENDARIO", "Standard", mvarSys)
   Else
      Call DockBarRightOf("BAR_CALENDARIO", "RH", mvarSys)
   End If
   mvarToolBar.Visible = True
   
   If Reload Then
      Call MontaTabSchedule(mvarIDLOJA, True)
      Call MontaDataProvider
      Call mvarMe_Resize
   End If
   Call PonteiroMeta
End Sub
Public Sub Hide()
   mvarMe.Hide
End Sub
Private Sub PonteiroMeta()
   Dim nAreaPrev0 As Integer
   Dim nAreaPrev As Integer
   Dim nAreaReal As Integer
   
   Dim nAreaRealProd As Integer
   Dim nAreaRealVend As Integer
   Dim nAreaPrevProd As Integer
   Dim nAreaPrevVend As Integer
   
   Dim Sql As String
   Dim nRs As Object
   Dim dData As Date
   Dim bExibeQtd As Boolean
      
   If gDebug Then MsgBox "TL_CALENDARIO.PonteiroMeta"
   If Not mvarPane.Visible Then Exit Sub

   If mvarSys.Propriedades("VIEWFXMETA") = "" Then
      mvarSys.Propriedades("VIEWFXMETA") = mvarSys.GetParam("VIEWFXMETA")
   End If
   bExibeQtd = (mvarSys.Propriedades("VIEWFXMETAQTD", True) = 1)

   mvarPane.PctPrev.Visible = (mvarSys.Propriedades("VIEWFXMETA") = 1)
   mvarPane.PctPrev0.Visible = mvarPane.PctPrev.Visible
   mvarPane.PctReal.Visible = mvarPane.PctPrev.Visible
   mvarPane.PrgGreen.Visible = mvarPane.PctPrev.Visible
   mvarPane.PrgRed.Visible = mvarPane.PctPrev.Visible
   mvarPane.PrgYellow.Visible = mvarPane.PctPrev.Visible
   If Not mvarPane.PctPrev.Visible Then Exit Sub
   
   Call DefineFaixa
   
   dData = CDate(mvarSys.xDb.Sysdate())
   
   
   Sql = " Select Count(*) [PREV]"
   Sql = Sql & " From OEVENTOAGENDA E"
   Sql = Sql & " Join OSERVICOEVT S On E.IDLOJA=S.IDLOJA and E.IDEVENTO=S.IDEVENTO"
   Sql = Sql & " Join OCLIENTE C On E.IDLOJA=C.IDLOJA And E.IDCLIENTE=C.IDCLIENTE And C.ISENTO=0"
   Sql = Sql & " Where E.STARTDATETIME >= Getdate()"
   Sql = Sql & " And Month(E.STARTDATETIME)=" & Month(dData)
   Sql = Sql & " And Year(E.STARTDATETIME)=" & Year(dData)
   Sql = Sql & " And S.IDTPSERVICO <>1"
   Sql = Sql & " And E.FLGCANCELADO<>1"
   Sql = Sql & " And E.IDLOJA=" & mvarIDLOJA
   If mvarSys.xDb.Abretabela(Sql, nRs) Then
      nAreaPrevProd = xVal(nRs("PREV") & "")
   End If
   
   Sql = "Select Count(*) [PROD]"
   Sql = Sql & " From OATENDIMENTO A "
   Sql = Sql & " Join OSESSAO S On S.IDLOJA=A.IDLOJA And S.IDATENDIMENTO=A.IDATENDIMENTO"
   Sql = Sql & " Join OCLIENTE C On C.IDLOJA=A.IDLOJA And C.IDCLIENTE=A.IDCLIENTE And C.ISENTO=0"
   Sql = Sql & " Where S.IDTPSERVICO <>1"
   Sql = Sql & " And Month(A.DTATEND) =" & Month(dData)
   Sql = Sql & " And Year(A.DTATEND) =" & Year(dData)
   Sql = Sql & " And A.IDLOJA=" & mvarIDLOJA
   If mvarSys.xDb.Abretabela(Sql, nRs) Then
      nAreaRealProd = xVal(nRs("PROD") & "")
   End If
   
   'Sql = "Select  Sum(I.QTDVENDA)-Sum(V.VLDESC/P.VLVENDA) as [Qtd]"
   Sql = "Select Sum(I.QTDVENDA*I.VLUNIT)/AVG(I.VLUNIT)- Sum(V.VLDESC)/AVG(I.VLUNIT) as [Qtd]"
   Sql = Sql & " From CITENSVENDA I"
   Sql = Sql & " Join CVENDA V On V.IDLOJA=I.IDLOJA And V.IDVENDA=I.IDVENDA"
   Sql = Sql & " Join OCLIENTE C On C.IDLOJA=V.IDLOJA And C.IDCLIENTE=V.IDCLIENTE And C.ISENTO=0"
   Sql = Sql & " Join SPRODUTO P On P.IDLOJA=I.IDLOJA And P.IDPROD=I.IDPROD"
   Sql = Sql & " Where Month(V.DTVENDA)=" & Month(dData)
   Sql = Sql & " And Year(V.DTVENDA)=" & Year(dData)
   Sql = Sql & " And V.IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And P.ESERVICO=1"
   Sql = Sql & " Group by Year(V.DTVENDA), Month(V.DTVENDA)"
   Sql = Sql & " Order by Year(V.DTVENDA), Month(V.DTVENDA)"
   If mvarSys.xDb.Abretabela(Sql, nRs) Then
      nAreaRealVend = xVal(nRs("Qtd") & "")
   End If
   If nAreaRealVend <> 0 Then
      nAreaPrev = nAreaPrevProd + IIf(nAreaRealVend < nAreaRealProd, nAreaRealVend, nAreaRealProd)
   Else
      nAreaPrev = nAreaPrevProd + nAreaRealProd
   End If
   nAreaReal = nAreaRealProd

   If nTot > 0 Then
      Call SetTag(mvarPane.PctReal, "QTD", nAreaReal)
      Call SetTag(mvarPane.PctReal, "PERC", CStr(CInt(nAreaReal / nTot * 100)) & "%")
      mvarPane.PctReal.ToolTipText = GetTag(mvarPane.PctReal, "PERC") & IIf(bExibeQtd, " / " & nAreaReal, "")
   End If
   mvarPane.PctReal.AutoSize = True
   mvarPane.PctReal.ZOrder 1
   mvarPane.PctReal.Top = mvarPane.PrgGreen.Top - mvarPane.PctReal.Height + 80
   If nTot > 0 Then mvarPane.PctReal.Left = nAreaReal * (mvarPane.Width / nTot) - 120
  
   nAreaPrev = IIf(nAreaPrev > nTot, nTot, nAreaPrev)
      
   If nTot > 0 Then
      Call SetTag(mvarPane.PctPrev, "QTD", nAreaPrev)
      Call SetTag(mvarPane.PctPrev, "PERC", CStr(CInt(nAreaPrev / nTot * 100)) & "%")
      mvarPane.PctPrev.ToolTipText = GetTag(mvarPane.PctPrev, "PERC") & IIf(bExibeQtd, " / " & nAreaPrev, "")
   End If
   mvarPane.PctPrev.AutoSize = True
   mvarPane.PctPrev.ZOrder 1
   mvarPane.PctPrev.Top = mvarPane.PctReal.Top
   If nTot > 0 Then mvarPane.PctPrev.Left = nAreaPrev * (mvarPane.Width / nTot) - 120
   
   nAreaPrev0 = xVal(mvarSys.Propriedades("VLPREV"))
   mvarPane.PctPrev0.Visible = (nAreaPrev0 <> 0)
   If nTot > 0 Then
      Call SetTag(mvarPane.PctPrev0, "QTD", nAreaPrev0)
      Call SetTag(mvarPane.PctPrev0, "PERC", CStr(CInt(nAreaPrev0 / nTot * 100)) & "%")
      mvarPane.PctPrev0.ToolTipText = GetTag(mvarPane.PctPrev0, "PERC") & IIf(bExibeQtd, " / " & nAreaPrev0, "")
   End If
   mvarPane.PctPrev0.AutoSize = True
   mvarPane.PctPrev0.ZOrder 1
   mvarPane.PctPrev0.Top = mvarPane.PctReal.Top
   If nTot > 0 Then mvarPane.PctPrev0.Left = nAreaPrev0 * (mvarPane.Width / nTot) - 120
        
   '*********
   '* Popular Tags
   If nTot > 0 Then Call SetTag(mvarPane.PctReal, "PERC", CStr(CInt(nAreaReal / nTot * 100)) & "%")
   Call SetTag(mvarPane.PctReal, "VALORPROD", nAreaRealProd)
   Call SetTag(mvarPane.PctReal, "VALORVENDA", nAreaRealVend)
   
   If nTot > 0 Then Call SetTag(mvarPane.PctPrev, "PERC", CStr(CInt(nAreaPrev / nTot * 100)) & "%")
   Call SetTag(mvarPane.PctPrev, "VALORPROD", nAreaPrevProd + nAreaRealProd)
   Call SetTag(mvarPane.PctPrev, "VALORVENDA", nAreaPrevProd + nAreaRealVend)
   
   If nTot > 0 Then Call SetTag(mvarPane.PctPrev0, "PERC", CStr(CInt(nAreaPrev0 / nTot * 100)) & "%")
   Call SetTag(mvarPane.PctPrev0, "VALOR", nAreaPrev0)
   
        
   If gDebug Then MsgBox "TL_CALENDARIO.PonteiroMeta End"
End Sub
Private Sub DefineFaixa()
   Dim TbMeta  As Object
   Set TbMeta = CriarObjeto("Banco_3R.TB_PMETA", False)
   With TbMeta
      Set .xDb = mvarSys.xDb
      If Not .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDMETA:=1) Then
         .IDLOJA = mvarIDLOJA
         .IDMETA = 1
         .FAIXA1 = 270
         .FAIXA2 = 320
         .FAIXA3 = .FAIXA2 + ((.FAIXA2 - .FAIXA1) / 2)
         .Salvar
      End If
   End With
   
   Fx1 = TbMeta.FAIXA1 ' 270
   Fx2 = TbMeta.FAIXA2 ' 320
   Fx3 = TbMeta.FAIXA3 ' Fx2 + ((Fx2 - Fx1) / 2) =345
   nTot = Fx3

   On Error Resume Next
   With mvarPane
      If nTot > 0 Then
         .PrgRed.Move 0, .Height - .PrgRed.Height, Fx1 * .Width / nTot, 100
         .PrgYellow.Move .PrgRed.Left + .PrgRed.Width, .PrgRed.Top, (Fx2 - Fx1) * .Width / nTot, .PrgRed.Height
         .PrgGreen.Move .PrgYellow.Left + .PrgYellow.Width, .PrgRed.Top, (Fx3 - Fx2) * .Width / nTot, .PrgRed.Height
      Else
         .PrgRed.Visible = False
         .PrgYellow.Visible = False
         .PrgGreen.Visible = False
         .PctPrev.Visible = False
         .PctPrev0.Visible = False
         .PctReal.Visible = False
      End If
   End With
End Sub
Private Sub InitTlEvento()
   Dim sKey As String
   sKey = "TLEVENTO"
   If xObjetos Is Nothing Then Set xObjetos = New Collection
   
   'If ExisteItem(xObjetos, sKey) Then
   '   Set mvarTLEvento = xObjetos(sKey)
   '   Call mvarTLEvento.Limpar
   'Else
      Set mvarTLEvento = Nothing
      Set mvarTLEvento = New TL_Evento
   'End If
   
   If Not mvarSys Is Nothing Then
      Set mvarTLEvento.Sys = mvarSys
      If Not mvarMe.CalendarControl Is Nothing Then
         Set mvarTLEvento.CalControl = mvarMe.CalendarControl
      End If
   End If

   If Not mvarTLEvento Is Nothing Then
      If Not ExisteItem(xObjetos, sKey) Then
         xObjetos.Add mvarTLEvento, sKey
      End If
   End If
End Sub
Private Sub InitTlAtend()
   Dim sKey As String
   sKey = "TLATENDIMENTO"
   If xObjetos Is Nothing Then Set xObjetos = New Collection
   
   'If ExisteItem(xObjetos, sKey) Then
   '   Set mvarTLAtend = xObjetos(sKey)
   '   Call mvarTLAtend.Limpar
   'Else
      Set mvarTLAtend = Nothing
      Set mvarTLAtend = New TL_Atendimento
   'End If
   
   If Not mvarSys Is Nothing Then
      Set mvarTLAtend.Sys = mvarSys
      If Not mvarMe.CalendarControl Is Nothing Then
         Set mvarTLAtend.CalControl = mvarMe.CalendarControl
      End If
   End If

   If Not mvarTLAtend Is Nothing Then
      If Not ExisteItem(xObjetos, sKey) Then
         xObjetos.Add mvarTLAtend, sKey
      End If
   End If
End Sub
Private Sub Class_Initialize()
   gDebug = False
   Set mvarMe = New FrmCalendario
   Set mvarPane = New FrmPaneCalendario
   Set gUnidades = New Collection
   
   InitTlEvento
   InitTlAtend
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Set mvarMe = Nothing
   Set mvarPane = Nothing
   Set gUnidades = Nothing
End Sub
Private Sub mvarCmmdBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   Dim nCtrl As Object
   Dim dIni As Date
   Dim dFim As Date
   Dim nViewType As CalendarViewType
   
   DoEvents
   Screen.MousePointer = vbHourglass
   
   dIni = mvarMe.CalendarControl.DayView.Selection.Begin
   dFim = mvarMe.CalendarControl.DayView.Selection.End
   nViewType = mvarMe.CalendarControl.ViewType
   
   mvarMe.CalendarControl.Enabled = False
   Select Case UCase(Control.Category)
      Case "CALENDARIO"
         Select Case Control.Id
            Case Mnu_CalLembrete:      Call ExibirLembrete
            Case Mnu_CalOpcao:         Call ExibirOpcao
            Case Mnu_CalOpcaoAvancada: Call ExibirOpcaoAvancada
            Case Mnu_CalEditOpen:      Call EditarEvento(mvarTLEvento.CalEvent)
            Case Mnu_CalEditDelete:    Call ExcluirEvento(mvarTLEvento.CalEvent)
            Case Mnu_CalEditCliente:   Call EditarCliente(mvarTLEvento.CalEvent)
            Case Mnu_CalEditVenda:     Call EditarVenda(mvarTLEvento.CalEvent)
            Case Mnu_CalEditConfirm:   Call ConfirmarEvento(Mnu_CalEditConfirm, mvarTLEvento.CalEvent, UCase(Control.Caption) = UCase("Desconfirmar"))
            Case Mnu_CalEditCancel:    Call ConfirmarEvento(Mnu_CalEditCancel, mvarTLEvento.CalEvent)
            Case Mnu_CalEditDelete2:   Call ExcluirAtendimento(mvarTLEvento.CalEvent)
            Case Mnu_CalNewEvent:      Call CriarEvento
            Case Mnu_CalTimeChange:    Call AlterarTimeZone
            Case Mnu_CalTime60, Mnu_CalTime30, Mnu_CalTime15, Mnu_CalTime10, Mnu_CalTime5:
               With mvarMe.CalendarControl
                  .ActiveView.ShowDay CDate(Format(.DayView.Selection.Begin, "dd/mm/yyyy"))
                  .DayView.TimeScale = xVal(Control.Caption)
                  .ViewType = nViewType
                  .DayView.Selection.Begin = dIni
                  .DayView.Selection.End = dFim
                  .DayView.ScrollToWorkDayBegin
                  .DayView.ScrollV .DayView.GetCellNumber(DateAdd("n", -1 * .DayView.TimeScale, .Options.WorkDayStartTime))
                  Call WriteIniFile(mvarSys.LocalReg, "FORMAT" & mvarSys.IDLOJA, "TimeScale", .DayView.TimeScale)
               End With
         End Select
      Case "BAR_CALENDARIO"
         For Each nCtrl In Control.Controls
            nCtrl.Checked = False
         Next
         mvarMe.CalendarControl.ActiveView.ShowDay CDate(Format(mvarMe.CalendarControl.DayView.Selection.Begin, "dd/mm/yyyy"))
         Select Case Control.Id
            Case Bar_CalDia:     mvarMe.CalendarControl.ViewType = xtpCalendarDayView: Control.Checked = True
            Case Bar_CalDiaUtil: mvarMe.CalendarControl.ViewType = xtpCalendarWorkWeekView: Control.Checked = True
            Case Bar_CalSemana:  mvarMe.CalendarControl.ViewType = xtpCalendarWeekView: Control.Checked = True
            Case Bar_CalMes:     mvarMe.CalendarControl.ViewType = xtpCalendarMonthView: Control.Checked = True
            'Me.CalendarControl.ActiveView.Cut
            'Me.CalendarControl.ActiveView.Copy
            'Me.CalendarControl.ActiveView.Paste
            Case Bar_CalPgSetup:       Call PageSetup
            Case Bar_CalPrintPreview:   Call PrintPreview
            Case Bar_CalPrint:         Call PrintCalendar
         End Select
         mvarMe.CalendarControl.DayView.Selection.Begin = dIni
         mvarMe.CalendarControl.DayView.Selection.End = dFim
         Call WriteIniFile(mvarSys.LocalReg, "FORMAT" & mvarSys.IDLOJA, "ViewType", mvarMe.CalendarControl.ViewType)
         Call UpdateToolbar
         
   End Select
   
   mvarMe.CalendarControl.Enabled = True
   Screen.MousePointer = vbDefault
End Sub
Private Sub UpdateToolbar()
    'Toolbar.Buttons(Bar_CalCut).Enabled = mvarMe.CalendarControl.ActiveView.CanCut
    'Toolbar.Buttons(Bar_CalCopy).Enabled = mvarMe.CalendarControl.ActiveView.CanCopy
    'Toolbar.Buttons(Bar_CalPaste).Enabled = mvarMe.CalendarControl.ActiveView.CanPaste
End Sub

Private Sub mvarCmmdBars_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
   If Control.Category = "CALENDARIO" Then
      If Control.Id = eMenuCal.Mnu_CalLembrete Then
         Control.Enabled = mvarMe.CalendarControl.IsRemindersEnabled
      End If
   End If
End Sub

Private Sub mvarMe_CalendarControlBeforeDrawDayViewCell(ByVal CellParams As XtremeCalendarControl.CalendarDayViewCellParams)
    If Weekday(CellParams.BeginTime) = 7 Then
      If TimeValue(CellParams.BeginTime) >= #6:00:00 PM# Then
         CellParams.BackgroundColor = vbRed ' RGB(255, 244, 188)
      End If
    End If
    
    ' standard colors are
    ' non-work cell Bk = RGB(255, 244, 188)
    '     work cell Bk = RGB(255, 255, 213)

   If CellParams.Selected Then
      'CellParams.BackgroundColor = RGB(20, 250, 50)
      Exit Sub
   End If
    
   If TimeValue(CellParams.BeginTime) >= #1:00:00 PM# And TimeValue(CellParams.BeginTime) < #2:00:00 PM# Then
      If Weekday(CellParams.BeginTime) <> 1 And Weekday(CellParams.BeginTime) <> 7 Then
         CellParams.BackgroundColor = RGB(198, 198, 198)
      End If
   End If

Exit Sub
'   Dim pTheme2007 As CalendarThemeOffice2007
'   Dim pCell As CalendarThemeDayViewCellParams
'
'   Set pTheme2007 = mvarMe.CalendarControl.Theme
'   If eObjType = xtpCalendarBeforeDraw_DayViewCell Then
'      Set pCell = DrawParams
'      If Not pCell.Selected Then
'         If TimeValue(pCell.BeginTime) >= CDate("12:00") And TimeValue(pCell.BeginTime) < CDate("14:00") And Weekday(pCell.BeginTime) <> 1 Then
'            pTheme2007.DayView.Day.Group.Cell.WorkCell.BackgroundColor = pTheme2007.DayView.Day.Group.Cell.NonWorkCell.BackgroundColor
'         End If
'      End If
'   End If
End Sub

Private Sub mvarMe_CalendarControlContextMenu(ByVal x As Single, ByVal y As Single)
    Dim Popup     As CommandBar
    Dim HitTest   As CalendarHitTestInfo
'    Dim TBAtend   As Object 'TB_OATENDIMENTO
'    Dim TbVenda   As Object 'TB_OATENDIMENTO_VENDA
    Dim vCtrl     As Variant
    Dim Sql       As String
    Dim MyRs      As Object
    
   Set HitTest = mvarMe.CalendarControl.ActiveView.HitTest
    
   
   If Not HitTest.ViewEvent Is Nothing Then
      Set mvarTLEvento.CalEvent = HitTest.ViewEvent.Event
      Set Popup = mvarSys.MDI.CommandBars.ContextMenus.Find(Pop_CalEdit)
      If Not Popup Is Nothing Then
         Popup.Controls.Find(xtpControlButton, Mnu_CalEditDelete2).Enabled = True
         If xVal(HitTest.ViewEvent.Event.CustomProperties("FLGCONFIRMADO")) = 1 Then
            Popup.Controls.Find(xtpControlButton, Mnu_CalEditConfirm).Caption = "Desconfirmar"
         Else
            Popup.Controls.Find(xtpControlButton, Mnu_CalEditConfirm).Caption = "Confirmar"
         End If
         
         For Each vCtrl In Popup.Controls
            vCtrl.Enabled = True
         Next
         
         Popup.Controls.Find(xtpControlButton, Mnu_CalEditVenda).Enabled = False
         Popup.Controls.Find(xtpControlButton, Mnu_CalEditOpen).Caption = "Abrir Evento"
         
         Sql = "Select IDATENDIMENTO"
         Sql = Sql & " From OATENDIMENTO "
         Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
         Sql = Sql & " And IDEVENTO=" & SqlNum(HitTest.ViewEvent.Event.Id)
         If mvarSys.xDb.Abretabela(Sql, MyRs) Then
            If MyRs.RecordCount > 0 Then
               Popup.Controls.Find(xtpControlButton, Mnu_CalEditOpen).Caption = "Abrir Atendimento" ' & ": " & MyRs("IDATENDIMENTO")
               Popup.Controls.Find(xtpControlButton, Mnu_CalEditOpen).DefaultItem = True
               Popup.Controls.Find(xtpControlButton, Mnu_CalEditDelete).Enabled = False
               Popup.Controls.Find(xtpControlButton, Mnu_CalEditCancel).Enabled = False
               Popup.Controls.Find(xtpControlButton, Mnu_CalEditConfirm).Enabled = False
               
               Sql = "Select IDVENDA"
               Sql = Sql & " From OATENDIMENTO_VENDA "
               Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
               Sql = Sql & " And IDATENDIMENTO =" & SqlNum(MyRs("IDATENDIMENTO") & "")
               If mvarSys.xDb.Abretabela(Sql, MyRs) Then
                  Popup.Controls.Find(xtpControlButton, Mnu_CalEditDelete2).Enabled = False
                  Popup.Controls.Find(xtpControlButton, Mnu_CalEditVenda).Enabled = True
                  If MyRs.RecordCount > 1 Then
                     Popup.Controls.Find(xtpControlButton, Mnu_CalEditVenda).Caption = "Vendas"
                  Else
                     Popup.Controls.Find(xtpControlButton, Mnu_CalEditVenda).Caption = "Venda: " & StrZero(MyRs("IDVENDA"), 6)
                  End If
               End If
            End If
         Else
            Popup.Controls.Find(xtpControlButton, Mnu_CalEditDelete2).Enabled = False
            Popup.Controls.Find(xtpControlButton, Mnu_CalEditVenda).Enabled = False
         End If
         Popup.ShowPopup
      End If
      Set mvarTLEvento.CalEvent = Nothing
      Call InitTlEvento
    
   ElseIf (HitTest.HitCode = xtpCalendarHitTestDayViewTimeScale) Then
      Set Popup = mvarSys.MDI.CommandBars.ContextMenus.Find(Pop_CalTime)
      If Not Popup Is Nothing Then
         Popup.ShowPopup
      End If
   Else
      Set Popup = mvarSys.MDI.CommandBars.ContextMenus.Find(Pop_CalNew)
      If Not Popup Is Nothing Then
         Popup.ShowPopup
      End If
    End If
End Sub
Private Sub mvarMe_CalendarControlDblClick()
   Dim HitTest As CalendarHitTestInfo
   
   Screen.MousePointer = vbHourglass
   mvarMe.CalendarControl.Enabled = False
   
   
   Set HitTest = mvarMe.CalendarControl.ActiveView.HitTest
   If HitTest.ViewEvent Is Nothing Then
      'Dim oTheme07   As CalendarThemeOffice2007
      'Set oTheme07 = mvarMe.CalendarControl.Theme
      'oTheme07.DayView.Day.Group.Cell.NonWorkCell.BackgroundColor = RGB(233, 235, 233)
      'mvarMe.CalendarControl.ActiveView.ShowDay CDate("04/02/2014")
       'RGB(224, 228, 224)
       'RGB(223, 229, 223)
      'RGB(230, 250, 216)
      'mvarMe.CalendarControl.Theme.RefreshMetrics
      'mvarMe.CalendarControl.Populate
      
      'mvarMe.CalendarControl.ActiveView.ShowDay CDate("04/02/2014")
      Call CriarEvento
   Else
      Call EditarEvento(HitTest.ViewEvent.Event)
   End If
    
'   Dim Events As CalendarEvents
'   If Not HitTest.HitCode = xtpCalendarHitTestUnknown Then
'      Set Events = mvarMe.CalendarControl.DataProvider.RetrieveDayEvents(HitTest.ViewDay.Date)
'   End If
   mvarMe.CalendarControl.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarMe_CalendarControlEventChanged(ByVal EventID As Long)
    Dim pEvent As CalendarEvent
    'Set pEvent = mvarMe.CalendarControl.DataProvider.GetEvent(EventID)
    
    ' pEvent Is Nothing for recurrence Ocurrence and Exception.
    ' See me.CalendarControl_PatternChanged for recurrence events changes.
    
    'If Not pEvent Is Nothing Then
       
    'End If

End Sub

Private Sub mvarMe_CalendarControlKeyDown(KeyCode As Integer, Shift As Integer)
   Dim BeginSelection   As Date
   Dim EndSelection     As Date
   Dim AllDay           As Boolean
   
   With mvarMe.CalendarControl.ActiveView
      If .GetSelection(BeginSelection, EndSelection, AllDay) Then
         If Shift = 2 Then
            Select Case KeyCode
               Case vbKeyC: If .CanCopy Then .Copy
               Case vbKeyX: If .CanCut Then .Cut
               Case vbKeyV: If .CanPaste Then .Paste
               Case vbKeyZ: If .CanUndo Then .Undo
            End Select
         Else
            If KeyCode = vbKeyReturn Or KeyCode = vbKeyF2 Then
               Call mvarMe_CalendarControlDblClick
            ElseIf KeyCode = vbKeyF5 Then
               Call RefreshCalendar
            End If
         End If
      End If
   End With
End Sub
Public Sub RefreshCalendar()
   Screen.MousePointer = vbHourglass
   Call MontaCalendario
   Call MontaTabSchedule(mvarIDLOJA, True)
   mvarMe.CalendarControl.DataProvider.ClearCache
   mvarMe.CalendarControl.Populate
   mvarPane.wndDatePicker.RedrawControl
   Screen.MousePointer = vbDefault
End Sub
              
Private Sub mvarMe_CalendarControlMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim HitTest As CalendarHitTestInfo
   
   If Button = 2 Then
      With mvarMe.CalendarControl
         Set HitTest = .ActiveView.HitTest
         If Not HitTest Is Nothing Then
            '.ActiveView.ResetSelection
            .ActiveView.Selection.Begin = HitTest.HitDateTime
            .ActiveView.Selection.End = HitTest.HitDateTime + .DayView.TimeScale * (1 / 24 / 60)
            .RedrawControl
         End If
      End With
   End If
End Sub

Private Sub mvarMe_CalendarControlOnReminders(ByVal Action As XtremeCalendarControl.CalendarRemindersAction, ByVal Reminder As XtremeCalendarControl.CalendarReminder)
    If g_bUseBuiltInCalendarDialogs Then
        g_dlgCalendarReminders.ShowRemindersWindow
        Exit Sub
    End If
    
    FrmLembrete.OnReminders Action, Reminder
    Set FrmLembrete.FrmCalendario = mvarMe
    
    If Action = xtpCalendarRemindersFire Then
        If ModalFormsRunningCounter = 0 Then
            If FrmLembrete.ctrlReminders.ListItems.Count > 0 Then
               FrmLembrete.Show vbModeless, mvarMe
            End If
        Else
            mvarMe.timerRMDForm.Enabled = True
        End If
    End If
End Sub
Private Sub mvarMe_Activate()
   If gDebug Then MsgBox "Activate Calendario"
   Static bCarregado As Boolean
   If bCarregado Then
      If Day(Now()) <> GetTag(mvarMe.CalendarControl, "DTACOES", 0) Then
         Call F_AcoesAgenda
      End If
   Else
      'Call MontarMenu
      'Call MontarToolbar
      Call MontaCalendario
      Call MontaTabSchedule(mvarIDLOJA, True)
      
      'Call ExibeAgendas
      bCarregado = True
      
      DoEvents
      If mvarMe.TabSchedule.Visible Then
         mvarMe.TabSchedule.Item(ReadIniFile(mvarSys.LocalReg, "FORMAT" & mvarSys.IDLOJA, "TabSala", 0)).Selected = True
      End If
      
'      Dim x As DatePickerDayMetrics
'      Dim xDia As CalendarViewEvent
'      Dim pEvents As CalendarEvents
'      'Call mvarPane_wndDatePickerDayMetrics(Now(), x)
'      Call mvarPane.wndDatePicker.ClearSelection
'      Call mvarPane.wndDatePicker.Select(CDate(Format(Now(), "DD/MM/YY")))
'      Call mvarPane.wndDatePicker.SetRange(CDate(Format(Now(), "DD/MM/YY")), CDate(Format(Now(), "DD/MM/YY")))
'      mvarMe.CalendarControl.ActiveView.Selection.Begin = CDate(Format(Now(), "DD/MM/YY"))
'      mvarMe.CalendarControl.ActiveView.Selection.End = CDate(Format(Now(), "DD/MM/YY"))
'
'      Set pEvents = mvarMe.CalendarControl.DataProvider.GetAllEventsRaw
'      Call NG_Cal.mvarCalendar_DoRetrieveDayEvents(CDate(Format(Now(), "DD/MM/YY")), pEvents)
'      Call mvarMe.CalendarControl.DataProvider.RetrieveDayEvents(CDate(Format(Now(), "DD/MM/YY")))
'
'      Call RefreshCalendar
      
      mvarMe.CalendarControl.Visible = True
      
   End If
   
'Static bcolor As Boolean
'bcolor = Not bcolor
'mvarMe.CalendarControl.DataProvider.LabelList.RemoveAll
'mvarMe.CalendarControl.DataProvider.LabelList.AddLabel 0, RGB(255, 255, 255), " "
'mvarMe.CalendarControl.DataProvider.LabelList.AddLabel 1, RGB(255, 231, 115), "Teste"
'If bcolor Then
'   mvarMe.CalendarControl.DataProvider.LabelList.AddLabel 2, RGB(165, 222, 99), "Depilação"
'   mvarSys.mdi.Caption = "Padrão"
'Else
'
'   mvarMe.CalendarControl.DataProvider.LabelList.AddLabel 2, RGB(132, 156, 231), "Depilação"
''   mvarMe.CalendarControl.DataProvider.LabelList.AddLabel 2, RGB(217, 238, 249), "Depilação"
''   mvarMe.CalendarControl.DataProvider.LabelList.AddLabel 2, RGB(140, 192, 108), "Depilação"
''   mvarMe.CalendarControl.DataProvider.LabelList.AddLabel 2, RGB(153, 204, 102), "Depilação"
''   mvarMe.CalendarControl.DataProvider.LabelList.AddLabel 2, RGB(179, 211, 155), "Depilação"
''   mvarMe.CalendarControl.DataProvider.LabelList.AddLabel 2, RGB(184, 210, 107), "Depilação"
''   mvarMe.CalendarControl.DataProvider.LabelList.AddLabel 2, RGB(209, 223, 156), "Depilação"
''   mvarMe.CalendarControl.DataProvider.LabelList.AddLabel 2, RGB(146, 192, 0), "Depilação"
'   mvarSys.mdi.Caption = "Novo"
'End If
'mvarMe.CalendarControl.DataProvider.LabelList.AddLabel 3, RGB(165, 206, 198), "Rejuvenescimento"
'mvarMe.CalendarControl.DataProvider.LabelList.AddLabel 4, RGB(214, 206, 132), "Acne"
'mvarMe.CalendarControl.DataProvider.LabelList.AddLabel 5, RGB(255, 148, 132), "Importante"
   
End Sub

Private Sub mvarMe_CalendarControlPrePopulateDay(ByVal ViewDay As XtremeCalendarControl.CalendarViewDay)
'    If me.CalendarControl.MultipleResources.Count > 1 Then
'
'        Dim arRes As New CalendarResources
'        'arRes.RemoveAll
'
'        If Day(ViewDay.Date) Mod 2 = 0 Then
'            arRes.Add me.CalendarControl.MultipleResources.Item(0)
'            ViewDay.SetMultipleResources arRes
'        ElseIf Day(ViewDay.Date) Mod 3 = 0 Then
'            arRes.Add me.CalendarControl.MultipleResources.Item(me.CalendarControl.MultipleResources.Count - 1)
'            ViewDay.SetMultipleResources arRes
'        End If
'    End If
End Sub

Private Sub mvarMe_CalendarControlSelectionChanged(ByVal SelType As XtremeCalendarControl.CalendarSelectionChanged)
'   If mvarMe.CalendarControl.ViewType <> SelType Then
'      mvarMe.CalendarControl.ViewType = SelType
'   End If
   
'   Select Case tbrPressed
'      Case Me.Toolbar.Buttons(1).Value
'         If Me.CalendarControl.ViewType <> xtpCalendarDayView Then
'            'Me.CalendarControl.ViewType = xtpCalendarDayView
'         End If
'      Case Me.Toolbar.Buttons(2).Value
'         If Me.CalendarControl.ViewType <> xtpCalendarWorkWeekView Then
'            Me.CalendarControl.ViewType = xtpCalendarWorkWeekView
'         End If
'      Case Me.Toolbar.Buttons(3).Value
'         If Me.CalendarControl.ViewType <> xtpCalendarWeekView Then
'            Me.CalendarControl.ViewType = xtpCalendarWeekView
'         End If
'      Case Me.Toolbar.Buttons(4).Value
'         If Me.CalendarControl.ViewType <> xtpCalendarMonthView Then
'            Me.CalendarControl.ViewType = xtpCalendarMonthView
'         End If
'   End Select
       
'    If SelType = xtpCalendarSelectionDays Then
'        Debug.Print "SelectionChanged. Day(s)."
'
'        If Me.CalendarControl.ActiveView.Selection.IsValid Then
'            Debug.Print Me.CalendarControl.ActiveView.Selection.Begin
'            Debug.Print Me.CalendarControl.ActiveView.Selection.End
'        End If
'    End If
'    If SelType = xtpCalendarSelectionEvents Then
'        Debug.Print "SelectionChanged. Event(s)."
'    End If
   Call UpdateToolbar
   Call PopulaStatusBar
End Sub
Private Sub PopulaStatusBar()
   Dim pEvent As CalendarEvent
   Dim oPane  As StatusBarPane
   
   On Error Resume Next
   
   Set oPane = mvarSys.MDI.CommandBars.StatusBar(3)
   oPane.Text = ""
   If mvarMe.CalendarControl.ActiveView.GetSelectedEvents.Count = 1 Then
      Set pEvent = mvarMe.CalendarControl.ActiveView.GetSelectedEvents(0).Event
      
      
      oPane.Text = pEvent.Subject & IIf(pEvent.Body = "", "", " - " & """" & pEvent.Body & """")
      'oPane.TextColor = vbBlack
      'oPane.BackgroundColor = 0
      'oPane.Alignment = xtpAlignmentLeft
      'oPane.Tooltip = pEvent.Subject
      'mvarSys.MDI.Refresh
   End If

Exit Sub
   Dim dData As Date
   Dim nTam  As Double
   Dim i     As Integer
   Dim Sql           As String
   
   Dim sCliTotal As String
   Dim sAreTotal As String
   Dim sCliOk As String
   Dim sAreOk As String
   Dim sCliNOk As String
   Dim sAreNOk As String
   Dim sCliTeste As String
   Dim sAreTeste As String
   
   dData = CDate(Format(mvarMe.CalendarControl.DayView.Selection.Begin, "dd/mm/yyyy"))
   If dData = CDate(GetTag(mvarMe.CalendarControl, "SELDATA", "01/01/0001")) Then
      Exit Sub
   End If
   Call SetTag(mvarMe.CalendarControl, "SELDATA", dData)
   
   sCliTotal = "0"
   sAreTotal = "0"
   sCliOk = "0"
   sAreOk = "0"
   sCliNOk = "0"
   sAreNOk = "0"
   sAreTeste = "0"
   sCliTeste = "0"
   
   '* Total dia
   Sql = ""
   Sql = Sql & " Select Count(Distinct E.IDCLIENTE) AS [CLI], COUNT(S.IDSERVICOEVT) AS [ARE]"
   Sql = Sql & " From OEVENTOAGENDA E"
   Sql = Sql & " Join OSERVICOEVT  S On E.IDLOJA=S.IDLOJA And E.IDEVENTO=S.IDEVENTO"
   Sql = Sql & " Where CONVERT(VARCHAR(10), E.STARTDATETIME, 103 )=" & SqlDate(dData, 1)
   Sql = Sql & " And E.IDLOJA=" & mvarIDLOJA
   If mvarSys.xDb.Abretabela(Sql) Then
      sCliTotal = mvarSys.xDb.RsAux("CLI")
      sAreTotal = mvarSys.xDb.RsAux("ARE")
   End If
   
   '* Atendimentos
   Sql = "Select Count(Distinct A.IDCLIENTE) AS [CLI], COUNT(S.IDSESSAO) AS [ARE]"
   Sql = Sql & " From OEVENTOAGENDA E"
   Sql = Sql & " Join OATENDIMENTO  A On A.IDLOJA=E.IDLOJA And A.IDEVENTO=E.IDEVENTO"
   Sql = Sql & " Join OSESSAO       S On A.IDLOJA=S.IDLOJA And A.IDATENDIMENTO=S.IDATENDIMENTO"
   Sql = Sql & " Where S.IDTPSERVICO <> 1"
   Sql = Sql & " And E.FLGCANCELADO<>1"
   Sql = Sql & " And CONVERT(VARCHAR(10), A.DTATEND, 103 )=" & SqlDate(dData, 1)
   Sql = Sql & " And E.IDLOJA=" & mvarIDLOJA
   If mvarSys.xDb.Abretabela(Sql) Then
      sCliOk = mvarSys.xDb.RsAux("CLI")
      sAreOk = mvarSys.xDb.RsAux("ARE")
   End If
   
   '* Cancelados
   Sql = ""
   Sql = Sql & " Select Count(Distinct E.IDCLIENTE) AS [CLI], COUNT(S.IDSERVICOEVT) AS [ARE]"
   Sql = Sql & " From OEVENTOAGENDA E"
   Sql = Sql & " Join OSERVICOEVT  S On E.IDLOJA=S.IDLOJA And E.IDEVENTO=S.IDEVENTO"
   Sql = Sql & " Where CONVERT(Varchar(10), E.STARTDATETIME, 103 )=" & SqlDate(dData, 1)
   Sql = Sql & " And E.FLGCANCELADO=1"
   Sql = Sql & " And S.IDTPSERVICO <> 1"
   Sql = Sql & " And E.IDLOJA=" & mvarIDLOJA
   If mvarSys.xDb.Abretabela(Sql) Then
      sCliNOk = mvarSys.xDb.RsAux("CLI")
      sAreNOk = mvarSys.xDb.RsAux("ARE")
   End If
   
   '* Teste
   Sql = ""
   Sql = Sql & "Select Count(Distinct A.IDCLIENTE) AS [CLITOTAL], COUNT(S.IDSESSAO) AS [AREATOTAL]"
   Sql = Sql & " From OATENDIMENTO A"
   Sql = Sql & " JOIN OSESSAO S On A.IDLOJA=S.IDLOJA AND A.IDATENDIMENTO=S.IDATENDIMENTO"
   Sql = Sql & " JOIN OEVENTOAGENDA E On A.IDLOJA=E.IDLOJA AND A.IDEVENTO=E.IDEVENTO"
   Sql = Sql & " Where S.IDTPSERVICO = 1"
   Sql = Sql & " And CONVERT(VARCHAR(10), A.DTATEND, 103 )=" & SqlDate(dData, 1)
   Sql = Sql & " And A.IDLOJA=" & mvarIDLOJA
   If mvarSys.xDb.Abretabela(Sql) Then
      sCliTeste = mvarSys.xDb.RsAux("CLITOTAL")
      sAreTeste = mvarSys.xDb.RsAux("AREATOTAL")
   End If
   
   
   With mvarMe.CommandBars1.StatusBar
      nTam = .Pane(0).Font.Size / 1.3
      
      .Pane(0).Text = "Data: " & dData
      .Pane(1).Text = "Clientes: " & sCliTotal
      .Pane(2).Text = "Áreas: " & sAreTotal
      .Pane(3).Text = "Atendimentos: " & sCliOk & " (C) / " & sAreOk & " (A)"
      .Pane(4).Text = "Cancelamentos: " & sCliNOk & " (C) / " & sAreNOk & " (A)"
      .Pane(5).Text = "Testes: " & sCliTeste & " (C)  / " & sAreTeste & " (A)"
      
      For i = 0 To .PaneCount - 1
         .Pane(i).Width = Len(.Pane(i).Text) * nTam
      Next
   End With
   
End Sub
Private Sub MontarStatusBar()
mvarMe.CommandBars1.DeleteAll
Exit Sub
   Dim StatusBar As XtremeCommandBars.IStatusBar
   Dim StPane As StatusBarPane
   Dim dData  As Date
   Dim nTam   As Double
   
   nTam = 1
   mvarMe.CommandBars1.DeleteAll
   Set StatusBar = mvarMe.CommandBars1.StatusBar
   With StatusBar
      .RemoveAll
      .DrawDisabledText = True
      .Visible = True
      
      Set StPane = .AddPane(.PaneCount)
      With StPane
         .Text = "Data: "
         .Alignment = xtpAlignmentCenter
         .Width = Len(.Text) * .Font.Size / nTam
         .Style = SBPS_DISABLED
      End With
      
      Set StPane = .AddPane(.PaneCount)
      With StPane
         .Text = "Clientes: "
         .Alignment = xtpAlignmentCenter
         .Width = Len(.Text) * .Font.Size / nTam
         .Style = SBPS_DISABLED
      End With
      
      Set StPane = .AddPane(.PaneCount)
      With StPane
         .Text = "Áreas: "
         .Alignment = xtpAlignmentCenter
         .Width = Len(.Text) * .Font.Size / nTam
         .Style = SBPS_DISABLED
      End With
      
      Set StPane = .AddPane(.PaneCount)
      With StPane
         .Text = "Atendimentos: "
         .Alignment = xtpAlignmentCenter
         .Width = Len(.Text) * .Font.Size / nTam
         .Style = SBPS_DISABLED
      End With
      
      Set StPane = .AddPane(.PaneCount)
      With StPane
         .Text = "Cancelamentos: "
         .Alignment = xtpAlignmentCenter
         .Width = Len(.Text) * .Font.Size / nTam
         .Style = SBPS_DISABLED
         '.Handle = Me.ProgressBar.hwnd
      End With
      
      Set StPane = .AddPane(.PaneCount)
      With StPane
         .Text = ""
         .Alignment = xtpAlignmentCenter
         .Width = Len(.Text) * .Font.Size / nTam
         .Style = SBPS_STRETCH + SBPS_DISABLED
      End With
   End With
End Sub
Private Sub mvarMe_CalendarControlViewChanged()
   Dim nCount  As Long
   Dim DtIni   As Date
   Dim DtFim   As Date
   Dim sAux    As String
   
   With mvarMe
      nCount = .CalendarControl.ActiveView.DaysCount
      DtIni = .CalendarControl.ActiveView.Days(0).Date
      DtFim = .CalendarControl.ActiveView.Days(nCount - 1).Date
       
      Select Case nCount
         Case 1: sAux = Format(DtIni, "Long Date")
         Case Is > 1
           'sAux = Format(DtIni, "Long Date") & " - " & Format(DtFim, "Long Date")
               
            sAux = StrZero(Day(DtIni), 2)
            If Year(DtIni) <> Year(DtFim) Or Month(DtIni) <> Month(DtFim) Then
               sAux = sAux & " de " & Format(DtIni, "mmmm")
               If Year(DtIni) <> Year(DtFim) Then
                  sAux = sAux & " de " & Year(DtIni)
               End If
            End If
            sAux = sAux & " - " & StrZero(Day(DtFim), 2)
            sAux = sAux & " de " & Format(DtFim, "mmmm")
            sAux = sAux & " de " & Year(DtFim)
            
      End Select
      .SccCalendario.Caption = sAux
   End With
   Call UpdateToolbar
End Sub

'============================================================================
'============================================================================
'== Default Background Color  Default Name Default LabelID                 ==
'== RGB(255,255,255)          None              0                          ==
'== RGB(255,148,132)          Important         1                          ==
'== RGB(132,156,231)          Business          2                          ==
'== RGB(165,222, 99)          Personal          3                          ==
'== RGB(231,231,214)          Vacation          4                          ==
'== RGB(255,181,115)          Must Attend       5                          ==
'== RGB(132,239,247)          Travel Required   6                          ==
'== RGB(214,206,132)          Needs Preparation 7                          ==
'== RGB(198,165,247)          Birthday          8                          ==
'== RGB(165,206,198)          Anniversary       9                          ==
'== RGB(255,231,115)          Phone Call        10                         ==
'============================================================================
'============================================================================
Private Sub mvarMe_Load()

   If gDebug Then MsgBox "Load Calendario"
   Call MontarTela
   Call MontarMenu
   Call MontarToolbar
   Call MontarStatusBar
   'Call MontaCalendario
   mvarMe.CalendarControl.Visible = False
   
   '=============================================================================
   DisableDragging_ForRecurrenceEvents = False
   DisableInPlaceCreateEvents_ForSaSu = False
   
   EnableScrollV_DayView = True
   EnableScrollH_DayView = True
   EnableScrollV_WeekView = True
   EnableScrollV_MonthView = True
   
   g_bUseBuiltInCalendarDialogs = False

   Set NG_Cal = New NG_Calendario
   Set NG_Cal.Sys = mvarSys
   Set NG_Cal.CalendarControl = mvarMe.CalendarControl
   Set NG_Cal.TlEvt = mvarTLEvento

   '************************
   '* Iniciar Ações de Agenda
   If gDebug Then MsgBox "AA"
   Call F_AcoesAgenda
   If gDebug Then MsgBox "Fim AA"

   '************************
   '* Define DataProvider
   Call MontaDataProvider
   
   mvarMe.TabSchedule.Visible = (mvarMe.TabSchedule.ItemCount > 1)
   Screen.MousePointer = vbDefault
   If gDebug Then MsgBox "Fim Load Calendario"
End Sub
Private Sub MontaDataProvider()
   Dim MyRs As Object
   Dim Sql As String
   Dim nCor As Long
   Dim nID  As Long
   Dim sTrat As String
   With mvarMe.CalendarControl

      .DataProvider.CacheMode = xtpCalendarDPCacheModeOff
      .SetDataProvider "Provider=custom;DSN=CalendarSQL_Server"    '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\events.mdb"
      .DataProvider.CacheMode = xtpCalendarDPCacheModeOnRepeat
      If Not .DataProvider.Open Then
         .DataProvider.Create
      End If
      Call ExibeAgendas
      If mvarSys.CODSIS = "P3R" Then
         Call AjustarLabelID
         With .DataProvider.LabelList
            If .Item(1).Name <> "Avaliação" Then
               .RemoveAll
               '.AddLabel 0, RGB(255, 255, 255), " "
               '.AddLabel 1001, RGB(255, 231, 115), "Avaliação"
               '.AddLabel 1002, RGB(217, 238, 249), "Dia Inteiro"
               '.AddLabel 9999, RGB(255, 148, 132), "Cancelado"
               
               .AddLabel 0, RGB(255, 255, 255), " "
               .AddLabel 1001, RGB(242, 254, 231), "Avaliação"
               .AddLabel 1002, RGB(217, 238, 249), "Dia Inteiro"
               .AddLabel 9999, RGB(255, 148, 132), "Cancelado"
               
               Sql = "Select IDTPTRATAMENTO, DSCTRATAMENTO, COR"
               Sql = Sql & " From OTPTRATAMENTO"
               Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
               'Sql = Sql & " And ATIVO=1"
               If mvarSys.xDb.Abretabela(Sql, MyRs) Then
                  While Not MyRs.EOF
                     nID = xVal(MyRs("IDTPTRATAMENTO") & "")
                     nCor = xVal(MyRs("COR") & "")
                     sTrat = MyRs("DSCTRATAMENTO")
                     If nCor = 0 Then
                        nCor = CorPadrao(sTrat)
                     End If
                     .AddLabel nID, nCor, sTrat
                     MyRs.MoveNext
                  Wend
                  
               Else
                  .AddLabel .Count, RGB(165, 222, 99), "Depilação"
                  .AddLabel .Count, RGB(165, 206, 198), "Rejuvenescimento"
                  .AddLabel .Count, RGB(214, 206, 132), "Acne"
               End If
            End If
         End With
      End If
   End With
   
   mvarMe.CalendarControl.Populate
   mvarPane.wndDatePicker.AutoSizeRowCol = True
   If mvarPane.wndDatePicker.RowCount < 2 Then
      mvarPane.wndDatePicker.AutoSizeRowCol = False
      mvarPane.wndDatePicker.ColumnCount = 1
      mvarPane.wndDatePicker.RowCount = 2
      mvarPane.wndDatePicker.Height = 4000
      mvarPane.wndDatePicker.Width = 2200
      mvarPane.wndDatePicker.Left = (mvarPane.SccCalendario.Width - mvarPane.wndDatePicker.Width) / 2
      mvarPane.wndDatePicker.Top = 980
   End If
   mvarPane.wndDatePicker.RedrawControl
   mvarMe.CalendarControl.Visible = mvarMe.Visible
   mvarMe.CalendarControl.RedrawControl
End Sub
Private Sub AjustarLabelID()
   Dim Sql As String
   Dim SqlW As String
   Dim MyRs As Object
   Dim bOk As Boolean
   Dim i As Integer
   Dim sAux As String
   Dim sXml As String
   Dim nPos As Integer
   Dim nPos2 As Integer
   
   Sql = "Select * "
   Sql = Sql & " From OEVENTOAGENDA"
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And LABELID>=1000"
   If Not mvarSys.xDb.ExisteReg(Sql) Then
      Sql = "Select E.IDLOJA, E.IDEVENTO, E.LABELID, E.FLGCANCELADO, E.ISALLDAYEVENT, E.CUSTOMPROPERTIESXMLDATA "
      Sql = Sql & ", S.IDTPSERVICO, S.IDTPTRATAMENTO"
      Sql = Sql & " From OEVENTOAGENDA E"
      Sql = Sql & " Left Join OSERVICOEVT S On E.IDLOJA=S.IDLOJA And E.IDEVENTO=S.IDEVENTO And (isnull(S.IDSERVICOEVT,0)=0 or S.IDSERVICOEVT=1)"
      Sql = Sql & " Where E.IDLOJA=" & mvarIDLOJA
      Sql = Sql & " Order By E.IDLOJA, E.IDEVENTO"
      If mvarSys.xDb.Abretabela(Sql, MyRs) Then
         While Not MyRs.EOF
            'If MyRs("IDEVENTO") = 1533 Then
            '   Sql = Sql
            '   sAux = MyRs("LABELID")
            'End If

            'SqlW = " Where IDLOJA=" & MyRs("IDLOJA")
            'SqlW = SqlW & " And IDEVENTO =" & MyRs("IDEVENTO")
            sAux = 0
            If MyRs("FLGCANCELADO") = 1 Then
               sAux = "9999"
            ElseIf MyRs("ISALLDAYEVENT") = 1 Then
               sAux = "1002"
            ElseIf xVal(MyRs("IDTPSERVICO") & "") = 1 Then
               sAux = "1001"
            ElseIf xVal(MyRs("IDTPSERVICO") & "") <= 0 Then
               sAux = "0"
            ElseIf xVal(MyRs("IDTPSERVICO") & "") > 1 Then
               sAux = xVal(MyRs("IDTPTRATAMENTO") & "")
            Else
               sAux = "0"
            End If
            sXml = MyRs("CustomPropertiesXMLData")
            nPos = InStr(sXml, "label0")
            If nPos > 0 Then
               nPos2 = InStr(186, sXml, "VariantType")
               If nPos2 > 0 Then
                  'sAux = Mid(sXml, nPos + 15, npos2 - nPos - 15 - 2)
                  sXml = Mid(MyRs("CustomPropertiesXMLData"), 1, nPos + 13) & """" & sAux & """" & Mid(MyRs("CustomPropertiesXMLData"), nPos2 - 1)
               End If
            End If
            
            Sql = "UPDATE OEVENTOAGENDA "
            Sql = Sql & " SET LABELID=" & sAux
            Sql = Sql & ",  CustomPropertiesXMLData='" & sXml & "'"
            Sql = Sql & ", ALTERSTAMP=1"
            Sql = Sql & ", TIMESTAMP=GetDate()"
            Sql = Sql & " Where IDLOJA=" & MyRs("IDLOJA")
            Sql = Sql & " And IDEVENTO =" & MyRs("IDEVENTO")
            bOk = mvarSys.xDb.Executa(Sql)
            If Not bOk Then i = i + 1
            MyRs.MoveNext
         Wend
      End If
   End If
End Sub
Public Sub F_AcoesAgenda()
  '**************
   '* Montar Tarefas de Sessões não Marcadas
   Dim Sql        As String
   Dim SqlProx    As String
   Dim NgTarefa   As Object
   Dim TBEvento   As Object
   Dim MyRs       As Object
   Dim Queries    As Collection
      
   Call SetTag(mvarMe.CalendarControl, "DTACOES", Day(Now))
      
'   Sql = "Select DISTINCT A.IDCLIENTE [Id.], C.NOME [Cliente], C.Tel1, C.Tel2, C.EMAIL, Convert(char, A.DTATEND, 103) [Data]"
'   Sql = Sql & ", 'Atendimento Sem Remarcação' [Obs]"
'   Sql = Sql & ", Convert(char, E.STARTDATETIME, 103) [Prox. Sessão]"
'   Sql = Sql & ", A.IDLOJA, A.IDEVENTO, A.DTATEND"
'   Sql = Sql & " From OATENDIMENTO A JOIN OCLIENTE C ON A.IDCLIENTE=C.IDCLIENTE"
'   Sql = Sql & " JOIN OSESSAO S ON A.IDLOJA=S.IDLOJA AND A.IDATENDIMENTO=S.IDATENDIMENTO AND S.IDTPSERVICO<>1"
'   Sql = Sql & " LEFT JOIN OEVENTOAGENDA  E On E.IDLOJA=A.IDLOJA AND E.IDCLIENTE=A.IDCLIENTE AND E.STARTDATETIME>A.DTATEND"
'   Sql = Sql & " LEFT JOIN OTAREFAEVT T On A.IDLOJA=T.IDLOJA AND A.IDEVENTO=T.IDEVENTO ANd T.IDTPTAREFA=" & eTPTArefa.TarNaoMarcado
'   Sql = Sql & " Where C.ATIVO = 1 And C.ISENTO = 0"
'   Sql = Sql & " AND  E.STARTDATETIME  is Null"
'   Sql = Sql & " AND  T.IDTPTAREFA  is Null "
'   Sql = Sql & " AND A.DTATEND > GetDate()-20"
'   Sql = Sql & " AND A.DTATEND < GetDate()-10"
'   Sql = Sql & " AND A.IDLOJA=" & mvarIDLOJA
'   Sql = Sql & " Order by A.DTATEND, 2, 1"
   
   SqlProx = "Select Min(PE.STARTDATETIME)"
   SqlProx = SqlProx & " From OEVENTOAGENDA PE"
   SqlProx = SqlProx & " Where PE.IDLOJA = A.IDLOJA"
   SqlProx = SqlProx & " And PE.IDCLIENTE=A.IDCLIENTE"
   SqlProx = SqlProx & " And PE.STARTDATETIME>A.DTATEND"
   
   Sql = "Select DISTINCT A.IDLOJA, A.IDCLIENTE [Id.], C.NOME [Cliente]" & vbNewLine
   Sql = Sql & ", C.Tel1, C.Tel2, C.EMAIL, Convert(char, A.DTATEND, 103) [Data], 'Atendimento Sem Remarcação' [Obs]" & vbNewLine
   Sql = Sql & ", A.IDATENDIMENTO, A.IDEVENTO, A.DTATEND" & vbNewLine
   Sql = Sql & ", (" & SqlProx & ") [Prox. Sessão]" & vbNewLine
   Sql = Sql & ", TR.DSCTRATAMENTO, TR.FREQUENCIA" & vbNewLine
   Sql = Sql & " From OATENDIMENTO A" & vbNewLine
   Sql = Sql & " Join OEVENTOAGENDA E On A.IDLOJA=E.IDLOJA And A.IDEVENTO=E.IDEVENTO" & vbNewLine
   Sql = Sql & " Join OCLIENTE C On A.IDLOJA=C.IDLOJA And A.IDCLIENTE=C.IDCLIENTE" & vbNewLine
   Sql = Sql & " Join OSESSAO S  On A.IDLOJA=S.IDLOJA And A.IDATENDIMENTO=S.IDATENDIMENTO And S.IDTPSERVICO<>1 And S.IDSESSAO=1" & vbNewLine
   Sql = Sql & " Join OTPTRATAMENTO TR On S.IDLOJA=TR.IDLOJA And S.IDTPTRATAMENTO=TR.IDTPTRATAMENTO And TR.FREQUENCIA>0" & vbNewLine
   Sql = Sql & " Left Join OTAREFAEVT T On A.IDLOJA=T.IDLOJA And A.IDEVENTO=T.IDEVENTO And T.IDTPTAREFA=" & eTpTarefa.TarNaoMarcado & vbNewLine
   Sql = Sql & " Where C.ATIVO = 1" & vbNewLine
   Sql = Sql & " And C.ISENTO = 0" & vbNewLine
   Sql = Sql & " And A.IDLOJA=" & mvarIDLOJA & vbNewLine
   Sql = Sql & " And  T.IDTPTAREFA  is Null" & vbNewLine
   Sql = Sql & " And A.DTATEND < GetDate()-20 And Year(A.DTATEND)>=2013" & vbNewLine
   Sql = Sql & " And (" & SqlProx & ") Is Null" & vbNewLine
   Sql = Sql & " And TR.FREQUENCIA>2"
   Sql = Sql & " Order by A.DTATEND, 3, 2"

If gDebug Then MsgBox "AA 01"
   If mvarSys.xDb.Abretabela(Sql, MyRs) Then
      Set Queries = New Collection
      Set NgTarefa = CriarObjeto("Tarefa3R.NG_TAREFA")
      Set NgTarefa.Sys = mvarSys
      
      Set TBEvento = CriarObjeto("BANCO_3R.TB_OEVENTOAGENDA")
      Set TBEvento.xDb = mvarSys.xDb
      While Not MyRs.EOF
         If TBEvento.Pesquisar(Ch_IDLOJA:=MyRs("IDLOJA"), Ch_IDEVENTO:=MyRs("IDEVENTO")) Then
            If gDebug Then MsgBox "AA 01.01"
            Call NgTarefa.SalvarTarefas(TBEvento, Queries, False, eTpTarefa.TarNaoMarcado)
         End If
         MyRs.MoveNext
      Wend
      If gDebug Then MsgBox "AA 01.02"
      If Queries.Count > 0 Then
         Call mvarSys.xDb.Executa(Queries)
      End If
      'Set NgTarefa = Nothing
   End If

'**************
'* Excluir Tarefas de Sessões já marcadas.
   
   
   '**************
   '* Grava Meta Prevista
   If NgTarefa Is Nothing Then
      Set NgTarefa = CriarObjeto("Tarefa3R.NG_TAREFA")
   End If
   Set NgTarefa.Sys = mvarSys
If gDebug Then MsgBox "AA 02"
   NgTarefa.GravaPrevisaoMeta
If gDebug Then MsgBox "AA saida"
End Sub

Private Sub MontarTela()
   mvarPane.wndDatePicker.AttachToCalendar mvarMe.CalendarControl
   With mvarMe
      .SccCalendario.VisualTheme = xtpShortcutThemeOffice2003
      .SccCalendario.GradientHorizontal = False
   End With
   Call MontarUnidades
   Call MontaTabSchedule(mvarIDLOJA)
End Sub
Public Sub MontarUnidades()
   Dim oUnid     As CLS_Unidade
   Dim Sql       As String
   Dim MyRs      As Object

   Sql = "Select * "
   Sql = Sql & " From OLOJA"
   Sql = Sql & " Where IDLOJA=" & mvarSys.IDLOJA
   Sql = Sql & " And IDCOLIGADA=" & mvarSys.IDCOLIGADA
   If Not mvarSys.xDb.Abretabela(Sql) Then
      Sql = "Select * "
      Sql = Sql & " From OLOJA"
      Sql = Sql & " Where IDLOJA=" & mvarSys.IDLOJA
      Sql = Sql & " Order By IDCOLIGADA"
      If mvarSys.xDb.Abretabela(Sql, MyRs) Then
         mvarSys.IDCOLIGADA = MyRs("IDCOLIGADA")
      End If
   End If
   Sql = "Select * "
   Sql = Sql & " From OLOJA"
   Sql = Sql & " Where IDCOLIGADA=" & mvarSys.IDCOLIGADA
   Sql = Sql & " Order By IDLOJA"
   If mvarSys.xDb.Abretabela(Sql, MyRs) Then
      While Not MyRs.EOF
         Set oUnid = New CLS_Unidade
         Set oUnid.Sys = mvarSys
         oUnid.IDLOJA = MyRs("IDLOJA")
         If Not ExisteItem(gUnidades, CStr(MyRs("IDLOJA"))) Then
            gUnidades.Add oUnid, CStr(MyRs("IDLOJA"))
         End If
         MyRs.MoveNext
         Set oUnid = Nothing
      Wend
   Else
      Sql = "Select * "
      Sql = Sql & " From OLOJA"
      Sql = Sql & " Where IDLOJA=" & mvarSys.IDLOJA
      Sql = Sql & " Order By IDLOJA"
      If mvarSys.xDb.Abretabela(Sql, MyRs) Then
         While Not MyRs.EOF
            Set oUnid = New CLS_Unidade
            Set oUnid.Sys = mvarSys
            oUnid.IDLOJA = MyRs("IDLOJA")
            gUnidades.Add oUnid, CStr(MyRs("IDLOJA"))
            mvarSys.IDCOLIGADA = MyRs("IDCOLIGADA")
            MyRs.MoveNext
            Set oUnid = Nothing
         Wend
      End If
      Sql = "Select * "
      Sql = Sql & " From OLOJA"
      Sql = Sql & " Where IDCOLIGADA=" & mvarSys.IDCOLIGADA
      Sql = Sql & " Order By IDLOJA"
      If mvarSys.xDb.Abretabela(Sql, MyRs) Then
         While Not MyRs.EOF
            Set oUnid = New CLS_Unidade
            Set oUnid.Sys = mvarSys
            oUnid.IDLOJA = MyRs("IDLOJA")
            If Not ExisteItem(gUnidades, CStr(MyRs("IDLOJA"))) Then
               gUnidades.Add oUnid, CStr(MyRs("IDLOJA"))
            End If
            MyRs.MoveNext
            Set oUnid = Nothing
         Wend
      End If
   End If
   
'   Dim TbUnidade As Object
'   Dim TbLoja    As Object
'   Dim TbSala    As Object
'
'   Set TbUnidade = CriarObjeto("BANCO_3R.TB_OLOJA", False)
'   Set gUnidades = New Collection
'   With TbUnidade
'      .xDb = mvarSys.xDb
'      If Not .Pesquisar(Ch_IDLOJA:=mvarSys.IDLOJA, Ch_Where:="IDCOLIGADA=" & mvarSys.IDCOLIGADA) Then
'         If .Pesquisar(Ch_IDLOJA:=mvarSys.IDLOJA, Ch_ORDERBY:="IDCOLIGADA") Then
'            mvarSys.IDCOLIGADA = .IDCOLIGADA
'         End If
'      End If
'
'      If .Pesquisar(Ch_Where:="IDCOLIGADA=" & mvarSys.IDCOLIGADA, Ch_ORDERBY:="IDLOJA") Then
'         While Not .Rs.EOF
'            .Popula
'            Set oUnid = New CLS_Unidade
'            Set oUnid.Sys = mvarSys
'            oUnid.IDLOJA = .IDLOJA
'            gUnidades.Add oUnid, CStr(.IDLOJA)
'            .Rs.MoveNext
'            Set oUnid = Nothing
'         Wend
'      Else
'         If .Pesquisar(Ch_IDLOJA:=mvarSys.IDLOJA, Ch_ORDERBY:="IDLOJA") Then
'            While Not .Rs.EOF
'               .Popula
'               Set oUnid = New CLS_Unidade
'               Set oUnid.Sys = mvarSys
'               oUnid.IDLOJA = .IDLOJA
'               gUnidades.Add oUnid, CStr(.IDLOJA)
'               mvarSys.IDCOLIGADA = .IDCOLIGADA
'               .Rs.MoveNext
'               Set oUnid = Nothing
'            Wend
'         End If
'         If .Pesquisar(Ch_Where:="IDCOLIGADA=" & mvarSys.IDCOLIGADA, Ch_ORDERBY:="IDLOJA") Then
'            While Not .Rs.EOF
'               .Popula
'               Set oUnid = New CLS_Unidade
'               Set oUnid.Sys = mvarSys
'               oUnid.IDLOJA = .IDLOJA
'               If Not ExisteItem(gUnidades, CStr(.IDLOJA)) Then
'                  gUnidades.Add oUnid, CStr(.IDLOJA)
'               End If
'               .Rs.MoveNext
'               Set oUnid = Nothing
'            Wend
'         End If
'      End If
'   End With
   Set mvarSys.Objetos("LOJAS") = gUnidades
End Sub
Private Sub MontaTabSchedule(pIDLOJA As Integer, Optional bRefresh As Boolean)
   Dim vSala As Variant
   Dim Item As TabControlItem
   
   Call SetTag(mvarMe, "MousePointer", Screen.MousePointer)
   Screen.MousePointer = vbHourglass
   
'   mvarMe.TabSchedule.Visible = False
   mvarMe.TabSchedule.RemoveAll
   If Not ExisteItem(gUnidades, CStr(pIDLOJA)) Or bRefresh Then
      Call MontarUnidades
      If Not ExisteItem(gUnidades, CStr(pIDLOJA)) Then Exit Sub
   End If
   For Each vSala In gUnidades(CStr(pIDLOJA)).Salas
      mvarMe.TabSchedule.InsertItem mvarMe.TabSchedule.ItemCount, vSala.NMSALA, 0, 0
      mvarMe.TabSchedule.Item(mvarMe.TabSchedule.ItemCount - 1).Tag = vSala.IDSALA
   Next
   mvarMe.TabSchedule.Width = (gUnidades(CStr(pIDLOJA)).Salas.Count + 1) * 930
   mvarMe.TabSchedule.Left = mvarMe.Width - mvarMe.TabSchedule.Width - 120
   mvarMe.TabSchedule.Height = 360

   If gUnidades(CStr(pIDLOJA)).Salas.Count > 1 Then
      mvarMe.TabSchedule.InsertItem mvarMe.TabSchedule.ItemCount, "Todos", 0, 0
      mvarMe.TabSchedule.Visible = True
   Else
      mvarMe.TabSchedule.RemoveAll
      mvarMe.TabSchedule.Visible = False
   End If
   If mvarMe.TabSchedule.Visible Then
      Set Item = mvarMe.TabSchedule.Selected
      For Each vSala In gUnidades(CStr(mvarIDLOJA)).Salas
         vSala.Selecionada = (Item.Tag = vSala.IDSALA Or UCase(Item.Caption) = "TODOS")
      Next
   End If
   Screen.MousePointer = GetTag(mvarMe, "MousePointer", vbDefault)
End Sub
Private Sub mvarMe_Resize()
   Dim nHeight As Long
   On Error Resume Next
   
   Dim l As Long, t As Long, r As Long, b As Long
   
   'Call MontaCalendario
   'Call AjustarLabelID
   With mvarMe
      nHeight = .Height - .CalendarControl.Top - 60  '5 * Screen.TwipsPerPixelY
      nHeight = IIf(nHeight < 0, 0, nHeight)
    
      '.CommandBars1.GetClientRect l, t, r, b
      'nHeight = b - .CalendarControl.Top
    
      '.SccCalendario.Move 0, 50, .ScaleWidth - 50, 420
      '.SccCalendario.Move 0, 420, .ScaleWidth - 50, 420
      .SccCalendario.Move 0, 60, .ScaleWidth - 80, 420
      .CalendarControl.Move 0, .SccCalendario.Top + .SccCalendario.Height, .SccCalendario.Width, nHeight
      
      
      .TabSchedule.Visible = False
      .TabSchedule.HeaderScrollOffset = False
      If .TabSchedule.ItemCount = 1 Then .TabSchedule.Width = 1370
      If .TabSchedule.ItemCount = 2 Then .TabSchedule.Width = 2280
      If .TabSchedule.ItemCount = 3 Then .TabSchedule.Width = 2760
      If .TabSchedule.ItemCount = 4 Then .TabSchedule.Width = 3660
      If .TabSchedule.ItemCount = 5 Then .TabSchedule.Width = 4960
      If .TabSchedule.ItemCount = 6 Then .TabSchedule.Width = 5460

      .TabSchedule.Left = .ScaleWidth - .TabSchedule.Width - 310
      .TabSchedule.Visible = (.TabSchedule.ItemCount > 2)
      
   End With
End Sub

Private Sub mvarMe_TabScheduleSelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   Dim i As Integer
   Dim vSala As Variant
   
   mvarMe.MousePointer = vbHourglass
   If Not ExisteItem(gUnidades, CStr(mvarIDLOJA)) Then Call MontarUnidades
   If Not ExisteItem(gUnidades, CStr(mvarIDLOJA)) Then Exit Sub
   If mvarMe.Visible Then
      For Each vSala In gUnidades(CStr(mvarIDLOJA)).Salas
         vSala.Selecionada = (Item.Tag = vSala.IDSALA Or UCase(Item.Caption) = "TODOS")
      Next
      Call MontaDataProvider
      Call WriteIniFile(mvarSys.LocalReg, "FORMAT" & mvarSys.IDLOJA, "TabSala", Item.Index)
   End If
   mvarMe.MousePointer = vbDefault
End Sub

Private Sub mvarMe_TimerCalTimer()
   Dim Sql As String
   Dim bGerou As Boolean
   Dim NG As Object
   Dim sPathBk As String
   Dim sBakName As String
   Dim bProducao  As Boolean
   
   Static dFile As Date
   Dim HoraFim As String
         
   bProducao = bProducao Or (UCase(mvarSys.xDb.Alias) = "PRODUCAO")
   bProducao = bProducao Or (UCase(mvarSys.xDb.Alias) = "PRODUCÃO")
   bProducao = bProducao Or (UCase(mvarSys.xDb.Alias) = "PRODUÇAO")
   bProducao = bProducao Or (UCase(mvarSys.xDb.Alias) = "PRODUÇÃO")
   If Not bProducao Then
      mvarMe.TimerCal.Enabled = False
      Exit Sub
   End If
   HoraFim = mvarSys.Propriedades("HORAFIM")
   If HoraFim <> "" Then
      If Day(HoraFim) <> Day(Now()) Then
         HoraFim = ""
         mvarSys.Propriedades("HORAFIM") = HoraFim
         bGerou = False
         Set NG = Nothing
      End If
   End If
   
   If HoraFim = "" Then
      Sql = "Select IsNull(Max(EndDateTime), GetDate()) As [Evento]" & vbNewLine
      Sql = Sql & " From OEVENTOAGENDA " & vbNewLine
      Sql = Sql & " Where Year(StartDateTime)=" & Year(mvarSys.xDb.Sysdate()) & vbNewLine
      Sql = Sql & " And Month(StartDateTime)=" & Month(mvarSys.xDb.Sysdate()) & vbNewLine
      Sql = Sql & " And Day(StartDateTime)=" & Day(mvarSys.xDb.Sysdate()) & vbNewLine
      If mvarSys.xDb.Abretabela(Sql) Then
         HoraFim = mvarSys.xDb.RsAux("Evento")
         mvarSys.Propriedades("HORAFIM") = HoraFim
      End If
   End If
   If (DateDiff("n", CDate(HoraFim), Now()) > 1 And DateDiff("n", dFile, Now()) > 20) Or (Hour(Now()) = 0 And Minute(Now()) > 40) Then
      sBakName = ""
      sBakName = Replace(mvarSys.Propriedades("NMLOJA"), " ", "")
      If InStr(sBakName, " ") <> 0 Then
         sBakName = Trim(Mid(sBakName, 1, InStr(sBakName, " ")))
      End If
      sBakName = mvarSys.Propriedades("IDCOLIGADA") & sBakName
   
      If sPathBk = "" Then sPathBk = ProcuraArquivo(Environ("programfiles") & "\Microsoft SQL Server\MSSQL\", sBakName & ".bak")
      If sPathBk = "" Then sPathBk = ProcuraArquivo(Environ("programfiles") & "\Microsoft SQL Server\MSSQL.1\", sBakName & ".bak")
      If sPathBk = "" Then sPathBk = ProcuraArquivo(Environ("programfiles") & "\Microsoft SQL Server\MSSQL.2\", sBakName & ".bak")
      If sPathBk = "" Then sPathBk = ProcuraArquivo(Environ("programfiles") & "\Microsoft SQL Server\MSSQL.3\", sBakName & ".bak")
      If ExisteArquivo(ResolvePathName(sPathBk) & sBakName & ".bak") Then
         dFile = FileDateTime(ResolvePathName(sPathBk) & sBakName & ".bak")
      Else
         dFile = DateAdd("n", Now(), -120)
      End If
      If DateDiff("n", dFile, Now()) < 20 Then
         bGerou = True
         Set NG = CriarObjeto("UTILITARIO3R.NG_UTILITARIO")
         With NG
            Set .Sys = mvarSys
            .F_BACKUP True, False
         End With
         If ExisteArquivo(ResolvePathName(sPathBk) & sBakName & ".bak") Then
            dFile = FileDateTime(ResolvePathName(sPathBk) & sBakName & ".bak")
         End If
         
      End If
   End If
End Sub

Private Sub mvarMe_Unload(Cancel As Integer)
  Dim bReminders As Boolean
  
  On Error Resume Next
   mvarMe.CalendarControl.DataProvider.Save
   mvarMe.CalendarControl.DataProvider.Close
   
   
   bReminders = mvarMe.CalendarControl.IsRemindersEnabled
   SaveSetting "PROJETO 3R", "AdvancedOptions", "RemindersEnabled", bReminders
   
   Unload FrmLembrete
End Sub

Private Sub mvarPane_Activate()
   If gDebug Then MsgBox "mvarPane_Activate"
   If GetTag(mvarPane, "1VEZ", "1") = 1 Then
      'Call mvarPane_Resize
      Call SetTag(mvarPane, "1VEZ", 0)
      mvarPane.SccCalendario.Visible = True
      mvarPane.SccCalendario2.Visible = True
      mvarPane.wndDatePicker.Visible = True
   End If
End Sub
Private Sub mvarPane_ChkAgendadoClick()
   If mvarPane.ChkAgendado.Value = vbUnchecked Then
      If mvarPane.ChkCancelados.Value = vbUnchecked Then
         mvarPane.ChkAgendado.Value = vbChecked
      End If
   End If
   Screen.MousePointer = vbHourglass
   NG_Cal.FLGAGENDADO = mvarPane.ChkAgendado.Value
   mvarMe.CalendarControl.DataProvider.ClearCache
   mvarMe.CalendarControl.Populate
   mvarPane.wndDatePicker.RedrawControl
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarPane_ChkCanceladosClick()
   If mvarPane.ChkCancelados.Value = vbUnchecked Then
      If mvarPane.ChkAgendado.Value = vbUnchecked Then
         mvarPane.ChkCancelados.Value = vbChecked
      End If
   End If
   NG_Cal.FLGCANCELADO = mvarPane.ChkCancelados.Value
   Screen.MousePointer = vbHourglass
   mvarMe.CalendarControl.DataProvider.ClearCache
   mvarMe.CalendarControl.Populate
   mvarPane.wndDatePicker.RedrawControl
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarPane_Load()
   If gDebug Then MsgBox "mvarPane_Load"
   With mvarPane.wndDatePicker
      .BorderStyle = xtpDatePickerBorderNone
      .ShowNoneButton = False
      .FirstDayOfWeek = 1
      .AskDayMetrics = True
      .BoldDaysPerIdleStep = 3
      .BoldDaysIdleStepTime_ms = 60
      .AutoSizeRowCol = True
      If .RowCount < 2 Then
         .AutoSizeRowCol = False
         .RowCount = 2
         .Height = 4000
         .Width = 2200
         .Left = 580
         .Top = 980
      End If
   End With
   Call CreateTaskPanel
   Call mvarPane_Resize
End Sub
Private Sub mvarPane_PctPrev0DblClick()
   Call PonteiroMeta
   Call ExibirInformacao(mvarPane.PctPrev0.ToolTipText, "Previsão Inicial")
End Sub
Private Sub mvarPane_PctPrev0MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Shift = 2 Then
      Call ExibirProducao
   End If
End Sub
Private Sub mvarPane_PctPrevDblClick()
   Call PonteiroMeta
   Call ExibirInformacao(mvarPane.PctPrev.ToolTipText, "Previsão")
End Sub
Private Sub mvarPane_PctPrevMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Shift = 2 Then
      Call ExibirProducao
   End If
End Sub
Private Sub mvarPane_PctRealDblClick()
   Call PonteiroMeta
   Call ExibirInformacao(mvarPane.PctReal.ToolTipText, "Real")
End Sub
Private Sub mvarPane_PctRealMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Shift = 2 Then
      Call ExibirProducao
   End If
End Sub
Private Sub mvarPane_Resize()
   On Error Resume Next
   With mvarPane
      .SccCalendario.Move 0, 0, .ScaleWidth, 420
      .SccCalendario2.Move 0, .SccCalendario.Height, .SccCalendario.Width, 285
      .wndDatePicker.Move 0, .SccCalendario2.Top + .SccCalendario2.Height + 50, .ScaleWidth, .ScaleHeight - 4000 '.wndDatePicker.Height
      '.wndDatePicker.BorderStyle = xtpDatePickerBorderOffice
      '.wndDatePicker.BorderStyle = 0
      '.wndDatePicker.Move -480, .SccCalendario2.Top + .SccCalendario2.Height + 50, 4250, 4500

      .wndDatePicker.AutoSizeRowCol = True
      If .wndDatePicker.RowCount < 2 Then
         .wndDatePicker.AutoSizeRowCol = False
         .wndDatePicker.RowCount = 2
         .wndDatePicker.Height = 4000
         .wndDatePicker.Width = 2200
         .wndDatePicker.Left = (.SccCalendario.Width - .wndDatePicker.Width) / 2
         .wndDatePicker.Top = 980
         If .wndDatePicker.RowCount <= 2 And .wndDatePicker.ColumnCount > 1 Then
            .wndDatePicker.Move 0, .SccCalendario2.Top + .SccCalendario2.Height + 50, .ScaleWidth, .ScaleHeight - 4000
            .wndDatePicker.AutoSizeRowCol = True
         End If
      End If
      .Picture1.Move 0, .wndDatePicker.Top + .wndDatePicker.Height, .ScaleWidth - 60, .ScaleHeight - .wndDatePicker.Height - .wndDatePicker.Top - 180
      '.Picture1.Height = .ScaleHeight - .wndDatePicker.Height - .wndDatePicker.Top - 180
      '.Picture1.Height = 3065
      '.Picture1.Visible = True
      '.Picture1.BackColor = vbRed
      .wndTaskPanel.Move 0, 0, .Picture1.Width, .Picture1.Height
      '.wndTaskPanel.Visible = True

'      .wndTaskPanel.Move 0, 0, .Picture1.ScaleWidth, .Picture1.ScaleHeight
'      .Picture1.Move 0, .SccCalendario2.Top + .SccCalendario2.Height, .ScaleWidth, .ScaleHeight - .SccCalendario.Height
'      .wndTaskPanel.Move 0, 0, .Picture1.ScaleWidth, .Picture1.ScaleHeight
      
      Call PonteiroMeta
   End With
   If mvarPane.Visible Then
      Call mvarPane_Activate
   End If
End Sub
Private Sub MontarContextMenu()
   Dim ContextMenu   As CommandBar
   Dim oMenuItem     As CommandBarControl
   
   '**************
   '* Pop-up Edit
   Set ContextMenu = mvarSys.MDI.CommandBars.ContextMenus.Add(Pop_CalEdit, "Context Menu Edit")
   With ContextMenu
      AddButtonCal .Controls, Mnu_CalEditOpen, "Abrir Evento"
      AddButtonCal .Controls, Mnu_CalEditDelete, "Excluir"
      AddButtonCal .Controls, Mnu_CalEditCliente, "Cliente", True
      AddButtonCal .Controls, Mnu_CalEditVenda, "Venda"
      AddButtonCal .Controls, Mnu_CalEditConfirm, "Confirmar", True
      AddButtonCal .Controls, Mnu_CalEditCancel, "Cancelar"
      AddButtonCal .Controls, Mnu_CalEditDelete2, "Excluir Atendimento", True
      
   End With
   ContextMenu.Controls.CreateOriginalControls
   
   '**************
   '* Pop-up New
   Set ContextMenu = mvarSys.MDI.CommandBars.ContextMenus.Add(Pop_CalNew, "Context Menu Edit")
   With ContextMenu
      AddButtonCal .Controls, Mnu_CalNewEvent, "Novo Compromisso"
   End With
   ContextMenu.Controls.CreateOriginalControls
      
   '**************
   '* Pop-up Time
   Set ContextMenu = mvarSys.MDI.CommandBars.ContextMenus.Add(Pop_CalTime, "Context Menu Edit")
   With ContextMenu
      AddButtonCal .Controls, Mnu_CalNewEvent, "Novo Compromisso"
      AddButtonCal .Controls, Mnu_CalTime60, "60", True
      AddButtonCal .Controls, Mnu_CalTime30, "30"
      AddButtonCal .Controls, Mnu_CalTime15, "15"
      AddButtonCal .Controls, Mnu_CalTime10, "10"
      AddButtonCal .Controls, Mnu_CalTime5, "5"
   End With
   ContextMenu.Controls.CreateOriginalControls
   
End Sub
Private Sub MontarToolbar()
   Dim oControl As CommandBarControl
   'Dim oToolBar As CommandBar
   
   'mvarSys.MDI.CommandBars.Icons = mvarMe.ImgToobar.Icons
   
   Set mvarToolBar = CriarToolbar(mvarSys, "BAR_CALENDARIO")
   If xVal(mvarSys.Propriedades("PONTO")) = 0 Then
      Call DockBarRightOf("BAR_CALENDARIO", "Standard", mvarSys)
   Else
      Call DockBarRightOf("BAR_CALENDARIO", "RH", mvarSys)
   End If
   
   With mvarToolBar
      .Customizable = False
      .Closeable = False
      .DefaultButtonStyle = xtpButtonIcon
      .Position = xtpBarTop
      .EnableDocking xtpFlagAlignAny
         
      Call CriarButtonToolbar(mvarToolBar, XTPControlType.xtpControlButton, Bar_CalDia, pCaption:="Dia", pIconId:=10002)
      Call CriarButtonToolbar(mvarToolBar, XTPControlType.xtpControlButton, Bar_CalDiaUtil, pCaption:="Dia Útil", pIconId:=10003, pChecked:=True)
      Call CriarButtonToolbar(mvarToolBar, XTPControlType.xtpControlButton, Bar_CalSemana, pCaption:="Semana", pIconId:=10004)
      Call CriarButtonToolbar(mvarToolBar, XTPControlType.xtpControlButton, Bar_CalMes, pCaption:="Mês", pIconId:=10005)
      Call CriarButtonToolbar(mvarToolBar, XTPControlType.xtpControlButton, Bar_CalPgSetup, pCaption:="Configurar Página", pIconId:=10006)
      Call CriarButtonToolbar(mvarToolBar, XTPControlType.xtpControlButton, Bar_CalPrintPreview, pCaption:="Visualizar Relatório", pBeginGroup:=True, pIconId:=10007)
      Call CriarButtonToolbar(mvarToolBar, XTPControlType.xtpControlButton, Bar_CalPrint, pCaption:="Imprimir", pIconId:=10008)
      'Set Control = .Controls.Add(XTPControlType.xtpControlButton, Bar_CalCut, "Cortar")
      'Set Control = .Controls.Add(XTPControlType.xtpControlButton, Bar_CalCopy, "Copiar")
      'Set Control = .Controls.Add(XTPControlType.xtpControlButton, Bar_CalPaste, "colar")
      
      Dim n As Object
      For Each n In .Controls
         n.Category = "BAR_CALENDARIO"
      Next
      
      .Position = 0
      .Visible = True
   End With
End Sub
Private Sub MontarMenu()
   Dim oToolBar      As CommandBars
   Dim oCommBar      As CommandBar
   Dim oMenuMain     As CommandBarControl
   Dim oMenuItem     As CommandBarControl
   Dim oMenuSubItem  As CommandBarControl
   Dim oMenuItemP    As CommandBarPopup
   Dim bMenuCarregado As Boolean
   
   On Error GoTo TrataErro
      
   Set oToolBar = mvarSys.MDI.CommandBars
   With oToolBar
      If .ActiveMenuBar Is Nothing Then
         Set oCommBar = .Add("Menu Calendario", xtpBarTop)
         oCommBar.Closeable = False
         oCommBar.Customizable = False
         oCommBar.EnableDocking xtpFlagHideWrap
      Else
         Set oCommBar = .ActiveMenuBar
      End If
      
      Set oMenuMain = oCommBar.FindControl(, Mnu_Calendario)
      If Not oMenuMain Is Nothing Then
         If (oMenuMain.Category = "CALENDARIO") Then
            oMenuMain.Visible = True
            GoTo Saida
         End If
      End If
      
      With oCommBar
         '************
         '* Menu Calendário
         Set oMenuMain = AddButtonCal(.Controls, Mnu_Calendario, "Calendário", False, xtpControlPopup)
         With oMenuMain.CommandBar
            AddButtonCal .Controls, Mnu_CalLembrete, "Lembretes"
            AddButtonCal .Controls, Mnu_CalOpcao, "Opções", True
            AddButtonCal .Controls, Mnu_CalOpcaoAvancada, "Opções Avançadas"
         End With
         Call MontarContextMenu
      End With
      .RecalcLayout
   End With
GoTo Saida

TrataErro:
   'Resume Next
   MsgBox "Erro em MontarMenu. Erro: " & Err.Number & "-" & Err.Description
Saida:
End Sub
Private Function AddButtonCal(Controls As CommandBarControls, Id As Long, Caption As String, _
                              Optional BeginGroup As Boolean = False, _
                              Optional ControlType As XTPControlType = xtpControlButton) As CommandBarControl
   
   Dim oMenuItem As CommandBarControl
    
   Set oMenuItem = Controls.Add(ControlType, Id, Caption)
   With oMenuItem
      .BeginGroup = BeginGroup
    
      .Category = "CALENDARIO"
      .Parameter = SetTag(.Parameter, "CARREGADO", 0)
      .Parameter = SetTag(.Parameter, "MENUCHILD", "S")
   End With
   
   Set AddButtonCal = oMenuItem
End Function
Private Sub CriarEvento()
   Dim BeginSelection   As Date
   Dim EndSelection     As Date
   Dim AllDay           As Boolean
   Dim oCalEvent        As CalendarEvent
       
   If g_bUseBuiltInCalendarDialogs Then
      Dim dlgCalendar As New CalendarDialogs
      dlgCalendar.ParentHWND = mvarMe.hwnd
      dlgCalendar.Calendar = mvarMe.CalendarControl
        
      dlgCalendar.ShowNewEvent
      Exit Sub
   End If
    
    
   Call mvarMe.CalendarControl.ActiveView.GetSelection(BeginSelection, EndSelection, AllDay)
   Set oCalEvent = mvarMe.CalendarControl.DataProvider.CreateEvent
    
   Call InitTlEvento
   'Set mvarTLEvento = New TL_Evento
   Set NG_Cal.TlEvt = mvarTLEvento
   'Set mvarTLEvento.CalControl = mvarMe.CalendarControl
   'Set mvarTLEvento.Sys = mvarSys
   Call mvarTLEvento.NewEvent(oCalEvent, BeginSelection, EndSelection, AllDay)
End Sub
Private Sub EditarEvento(ByRef ModEvent As CalendarEvent)
        
    If g_bUseBuiltInCalendarDialogs Then
        Dim dlgCalendar As New CalendarDialogs
        dlgCalendar.ParentHWND = mvarMe.hwnd
        dlgCalendar.Calendar = mvarMe.CalendarControl
        
        dlgCalendar.ShowEditEvent ModEvent
        
        Exit Sub
    End If
        
    If ModEvent.RecurrenceState <> xtpCalendarRecurrenceNotRecurring Then
        FrmSelItemSerie.m_bOcurrence = True
        FrmSelItemSerie.m_bDeleteRequest = False
        FrmSelItemSerie.m_strEventSubject = ModEvent.Subject
        
        FrmSelItemSerie.Show vbModal
        
        If FrmSelItemSerie.m_bOK = False Then
            Exit Sub
        ElseIf FrmSelItemSerie.m_bOcurrence Then
            If ModEvent.RecurrenceState = xtpCalendarRecurrenceOccurrence Then
                Set ModEvent = ModEvent.CloneEvent
                ModEvent.MakeAsRException
            End If
        Else
            ' Series
            Set ModEvent = ModEvent.RecurrencePattern.MasterEvent
        End If
    End If
    
   
   If ExisteAtendimento(ModEvent.Id) <> 0 Then
      Call InitTlAtend
      Set mvarTLAtend.CalEvent = ModEvent
      mvarTLAtend.IDLOJA = mvarIDLOJA
      mvarTLAtend.IDEVENTO = ModEvent.Id
      mvarTLAtend.IDSALA = ModEvent.ScheduleID
      mvarTLAtend.Show vbModal
   Else
      Call InitTlEvento
      Set NG_Cal.TlEvt = mvarTLEvento
      Set mvarTLEvento.CalEvent = ModEvent
      mvarTLEvento.IDEVENTO = ModEvent.Id
      If mvarTLEvento.isLoad Then
         mvarTLEvento.PopulaTela
      End If
      mvarTLEvento.Show
   End If
End Sub
Private Sub EditarCliente(ByRef ModEvent As CalendarEvent)
   Dim MyCliente As Object
   
   If ModEvent Is Nothing Then Exit Sub
   If xVal(ModEvent.CustomProperties("IDCLIENTE")) <= 0 Then Exit Sub
   
   Call ShowCliente(mvarSys, mvarIDLOJA, xVal(ModEvent.CustomProperties("IDCLIENTE")))
End Sub
Private Sub EditarVenda(ByRef ModEvent As CalendarEvent)
   Dim TlVenda As Object
   Dim MyRs As Object
   Dim Sql  As String
   
   If ModEvent Is Nothing Then Exit Sub
   If xVal(ModEvent.Id) <= 0 Then Exit Sub
   
   Sql = "Select V.IDLOJA, V.IDVENDA, A.IDATENDIMENTO, A.IDCLIENTE"
   Sql = Sql & " From OATENDIMENTO_VENDA V"
   Sql = Sql & " Join OATENDIMENTO A On A.IDLOJA=V.IDLOJA And A.IDATENDIMENTO=V.IDATENDIMENTO "
   Sql = Sql & " Where A.IDLOJA=" & SqlNum(mvarIDLOJA)
   Sql = Sql & " And A.IDEVENTO=" & SqlNum(ModEvent.Id)
   If mvarSys.xDb.Abretabela(Sql, MyRs) Then
      Set TlVenda = CriarObjeto("Financ3R.TL_Venda")
      With TlVenda
         Set .Sys = mvarSys
         Set .CalControl = mvarMe.CalendarControl
         
         .IDLOJA = MyRs("IDLOJA")
         .IDATENDIMENTO = MyRs("IDATENDIMENTO")
         .IDCLIENTE = MyRs("IDCLIENTE")
         .IDVENDA = MyRs("IDVENDA")
         
         .Show vbModal
      End With
      Set TlVenda = Nothing
   End If
End Sub
Private Function ExisteAtendimento(pIDEVENTO As Long) As Long
   Dim Sql     As String
   
   Sql = "Select IDATENDIMENTO " & vbNewLine
   Sql = Sql & " From OATENDIMENTO" & vbNewLine
   Sql = Sql & " Where IDLOJA = " & mvarIDLOJA & vbNewLine
   Sql = Sql & " And IDEVENTO = " & pIDEVENTO & vbNewLine
   Sql = Sql & " And IDATENDIMENTO<>0"
   If mvarSys.xDb.ExisteReg(Sql) Then
      ExisteAtendimento = mvarSys.xDb.RsAux("IDATENDIMENTO")
   End If
End Function
Private Sub ExcluirAtendimento(ByRef ModEvent As CalendarEvent)
   Dim bExiste As Boolean
   Dim oEvt As CalendarEvent
   
   Dim NgCal As NG_Calendario
   Set NgCal = New NG_Calendario
   Set NgCal.Sys = mvarSys

   If Not NgCal.VerificaEvento(ModEvent, False) Then GoTo Saida
   
   Call NgCal.ExcluiAtendimento(ModEvent.Id)
   
   Call RefreshEvent(mvarSys, mvarMe.CalendarControl, ModEvent)
Saida:

End Sub
Private Sub ConfirmarEvento(pMode As eMenuCal, ByRef ModEvent As CalendarEvent, Optional bDesconfirmar As Boolean)
   Dim TBEvento As Object 'TB_OEVENTOAGENDA
   Dim MyTbServ As Object 'TB_OSERVICOEVT
   Dim bAchou   As Boolean
   Dim nIDTRAT  As Integer
   Dim nIDSERV  As Integer
   Dim nLabel0  As Integer
   Dim pMotivo  As String
   
   If ModEvent Is Nothing Then Exit Sub
   
   Set TBEvento = CriarObjeto("BANCO_3R.TB_OEVENTOAGENDA")
   With TBEvento
      .xDb = mvarSys.xDb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEVENTO:=ModEvent.Id) Then
         nLabel0 = .LabelID
      
         Set MyTbServ = CriarObjeto("BANCO_3R.TB_OSERVICOEVT")
         Set MyTbServ.xDb = mvarSys.xDb
         bAchou = MyTbServ.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEVENTO:=ModEvent.Id)
         If bAchou Then
            nIDTRAT = MyTbServ.IDTPTRATAMENTO
            nIDSERV = MyTbServ.IDTPSERVICO
         Else
            nIDTRAT = 0
            nIDSERV = 0
         End If
         Select Case pMode
            Case Mnu_CalEditConfirm:
               .FLGCONFIRMADO = IIf(bDesconfirmar, 0, 1)
               .FLGCANCELADO = 0
               .LabelID = DefineLabelID(nIDTRAT, nIDSERV, .FLGCANCELADO)
               ModEvent.CustomProperties("label0") = .LabelID
               .CustomPropertiesXMLData = ModEvent.CustomProperties.SaveToString
               
            Case Mnu_CalEditCancel:
               .FLGCONFIRMADO = 0
               .FLGCANCELADO = 1
               .LabelID = DefineLabelID(nIDTRAT, nIDSERV, .FLGCANCELADO)
               ModEvent.CustomProperties("label0") = .LabelID
               .CustomPropertiesXMLData = ModEvent.CustomProperties.SaveToString
               
               TBEvento.Body = ExibirMotivo(TBEvento.Body)
               
        End Select
         If .isDirt Then
            If .Salvar Then
               ModEvent.Label = .LabelID
               ModEvent.Body = TBEvento.Body
               ModEvent.CustomProperties("FLGCANCELADO") = .FLGCANCELADO
               ModEvent.CustomProperties("FLGCONFIRMADO") = .FLGCONFIRMADO
               ModEvent.CustomProperties("label0") = .LabelID
               If Not NG_Cal Is Nothing Then
                  On Error Resume Next
                  'NG_Cal.IDCLIENTE = .IDCLIENTE
                  'If Not NG_Cal.TlEvt.Form.isLoad Is Nothing Then
                     'If Not NG_Cal.TlEvt.Form.Visible Is Nothing Then
                        If NG_Cal.TlEvt.CalControl Is Nothing Then
                           Set NG_Cal.TlEvt.CalControl = mvarMe.CalendarControl
                        End If
                        NG_Cal.TlEvt.Form.ChkFLGCONFIRMADO.Value = .FLGCONFIRMADO
                        NG_Cal.TlEvt.Form.ChkFLGCANCELADO.Value = .FLGCANCELADO
                        NG_Cal.TlEvt.Form.ChkMeeting.Value = .ISMEETING
                        mvarTLEvento.Form.ChkFLGCONFIRMADO.Value = .FLGCONFIRMADO
                        mvarTLEvento.Form.ChkFLGCANCELADO.Value = .FLGCANCELADO
                        mvarTLEvento.Form.ChkMeeting.Value = .ISMEETING
                     'End If
                  'End If
               End If
            Else
               ModEvent.Label = nLabel0
               ModEvent.CustomProperties("label0") = nLabel0
            End If
            Call mvarMe.CalendarControl.DataProvider.ChangeEvent(ModEvent)
      End If
     End If
   End With
    Set TBEvento = Nothing
End Sub
Private Function ExibirMotivo(pTexto As String) As String
   Dim MyEditor As Object
   Set MyEditor = CriarObjeto("DSACTIVE.EDITOR")
   With MyEditor
      .FundoTela = Empty
      .BtnPrint = False
      '.CabDetalhe = "Cab detahle"
      '.TituloRel = "Tit"
      
      .Caption = "Qual o motivo do Cancelamento?"
      .Texto = pTexto
      .Width = 5000 'mvarMe.Width
      .Height = 3500 'mvarMe.Height
      .Edicao = True
      '.SetIcon mvarMe
   
      .Show
      
      pTexto = .Texto
   End With
   Set MyEditor = Nothing
   ExibirMotivo = pTexto
End Function
Private Sub ExcluirEvento(ByRef ModEvent As CalendarEvent)
   Dim bExiste As Boolean
   
   Dim NgCal As NG_Calendario
   Set NgCal = New NG_Calendario
   Set NgCal.Sys = mvarSys
   If Not NgCal.VerificaEvento(mvarTLEvento.CalEvent, True) Then GoTo Saida
   
   bExiste = True
   With ModEvent
      'xtpCalendarRecurrenceOccurrence=2 / xtpCalendarRecurrenceException=3
      If .RecurrenceState = 2 Or .RecurrenceState = 3 Then
         FrmSelItemSerie.m_bOcurrence = True
         FrmSelItemSerie.m_bDeleteRequest = True
         FrmSelItemSerie.m_strEventSubject = mvarTLEvento.CalEvent.Subject
        
         FrmSelItemSerie.Show vbModal
        
         If FrmSelItemSerie.m_bOK Then
            '**************
            '* Exclui Série
            If Not FrmSelItemSerie.m_bOcurrence Then
               Call mvarMe.CalendarControl.DataProvider.DeleteEvent(.RecurrencePattern.MasterEvent)
               bExiste = False
            End If
         Else
            Exit Sub
         End If
      End If
   End With
   
   If bExiste Then
      mvarMe.CalendarControl.DataProvider.DeleteEvent mvarTLEvento.CalEvent
   End If
   mvarMe.CalendarControl.Populate
   'mvarCalControl.RedrawControl
Saida:

End Sub
Private Sub ExibirLembrete()
   If g_bUseBuiltInCalendarDialogs Then
      g_dlgCalendarReminders.ShowRemindersWindow
      Exit Sub
   End If
   If ModalFormsRunningCounter = 0 Then
      If Not FrmLembrete.Visible Then
         FrmLembrete.OnReminders xtpCalendarRemindersFire, Nothing
      End If
      If FrmLembrete.ctrlReminders.ListItems.Count > 0 Then
         FrmLembrete.Show vbModeless, mvarMe
      End If
   End If
End Sub
Private Sub AlterarTimeZone()
   Dim dlgCalendar As New CalendarDialogs
   
   If g_bUseBuiltInCalendarDialogs Then
      dlgCalendar.ParentHWND = mvarMe.hwnd
      dlgCalendar.Calendar = mvarMe.CalendarControl
      dlgCalendar.ShowTimeScaleProperties
      Exit Sub
   End If
   frmTimeZone.Show vbModal, Me
End Sub
Private Sub ExibirOpcao()
   Set FrmOpcao.mvarCalControl = mvarMe.CalendarControl
   FrmOpcao.Show vbModal
   
   Call mvarSys.SaveParam(pCODPARAM:="WorkDayStartTime", pVLPARAM:=mvarMe.CalendarControl.Options.WorkDayStartTime)
   Call mvarSys.SaveParam(pCODPARAM:="WorkDayEndTime", pVLPARAM:=mvarMe.CalendarControl.Options.WorkDayEndTime)
End Sub
Private Sub ExibirOpcaoAvancada()
   Set FrmOpcaoAvancada.mvarCalControl = mvarMe.CalendarControl
   FrmOpcaoAvancada.Show vbModal
End Sub
Private Sub MontaCalendario()
   Dim dToday     As Date
   Dim oTheme07   As CalendarThemeOffice2007
   Dim bLoadXml   As Boolean
   
   Call SetTag(mvarMe, "MousePointer", Screen.MousePointer)
   Screen.MousePointer = vbHourglass
   
   With mvarMe.CalendarControl
      '.SetTheme (New CalendarThemeOffice2007)
      Set oTheme07 = New CalendarThemeOffice2007
      
      ' Load customized theme options
      Dim px As PropExchange
      bLoadXml = False
      Set px = XtremeCalendarControl.CreatePropExchange
      If px.CreateAsXML(True, "CalendarThemeOffice2007") Then
         If px.LoadFromFile(mvarSys.ExePath & "\cfgCalendarThemeOffice2007.xml") Then
            .SetTheme oTheme07
            .Theme.DoPropExchange px
            bLoadXml = True
         End If
      End If
      
      With oTheme07
         .DayView.Day.Group.Cell.NonWorkCell.BackgroundColor = RGB(233, 235, 233)
         If Not bLoadXml Then
            .BaseColor = mvarMe.SccCalendario.GradientColorLight
            With .DayView.TimeScale
               .BackgroundColor = mvarMe.SccCalendario.GradientColorLight
               .LineColor = oTheme07.WeekView.Day.BackgroundDarkColor
               .TimeTextBigBase.Color = &H404040       'vbBlack
               .TimeTextSmall.Color = &H404040       'vbBlack
               
               .HeightFormula.Multiplier = 17
            End With
         End If
         ' XXX CRIAR EVENTO E ASSOCIAR AOS ICONES
         .CustomIcons = mvarMe.ImageEvento.Icons
      End With
      .SetTheme oTheme07
'      .SetTheme Nothing
'      .Populate
     
      If Not NG_Cal Is Nothing Then NG_Cal.SaveTheme
      
      If Not .BeforeDrawThemeObjectFlags Is Nothing Then
         '.BeforeDrawThemeObjectFlags.SetFlag (CalendarBeforeDrawThemeObject.xtpCalendarBeforeDraw_DayViewDay)
         '.BeforeDrawThemeObjectFlags.SetFlag -1
         '.BeforeDrawThemeObjectFlags.SetFlag CalendarBeforeDrawThemeObject.xtpCalendarBeforeDraw_DayViewTimeScaleCell
         'CalendarBeforeDrawThemeObject.xtpCalendarBeforeDraw_DayViewCell
      End If
      
      .EnableReminders GetSetting("PROJETO 3R", "AdvancedOptions", "RemindersManEnabled", True)
      With .Options
         .WorkDayStartTime = mvarSys.GetParam(pCODPARAM:="WorkDayStartTime", Default:="09:00:00", pDescricao:="Início do Dia Útil", pSECAO:="CALENDARIO")
         .WorkDayEndTime = mvarSys.GetParam(pCODPARAM:="WorkDayEndTime", Default:="18:00:00", pDescricao:="Fim do Dia Útil", pSECAO:="CALENDARIO")
         '.GetParam(pCODPARAM:="NFE", Default:="0", pDescricao:="Nota Fiscal Eletrônica", pSECAO:="NFE")
         .WorkWeekMask = xtpCalendarDayMo_Fr Or xtpCalendarDaySaturday
         .FirstDayOfTheWeek = xtpCalendarDaySunday
         '.UseOutlookFontGlyphs = True
         .EnableInPlaceCreateEvent = False
         .EnableInPlaceEditEventSubject_AfterEventResize = False
         .EnableInPlaceEditEventSubject_ByF2 = False
         .EnableInPlaceEditEventSubject_ByMouseClick = False
         .EnableInPlaceEditEventSubject_ByTab = False
           
      End With
'      dToday = Now
'      .ViewType = xtpCalendarDayView
'      .DayView.ShowDays dToday - 2, dToday + 2
'      .DayView.ScrollV 8 * (60 / .DayView.TimeScale)
      .ActiveView.ShowDay Now()
      .ViewType = ReadIniFile(mvarSys.LocalReg, "FORMAT" & mvarSys.IDLOJA, "ViewType", xtpCalendarWorkWeekView)
      .DayView.TimeScale = ReadIniFile(mvarSys.LocalReg, "FORMAT" & mvarSys.IDLOJA, "TimeScale", 30)
      .DayView.ScrollToWorkDayBegin
      .DayView.ScrollV .DayView.GetCellNumber(DateAdd("n", -1 * .DayView.TimeScale, .Options.WorkDayStartTime))
      .DayView.MinColumnWidth = 60
      .Visible = True
      .Populate
      .RedrawControl
      
      If g_bUseBuiltInCalendarDialogs Then
         Set g_dlgCalendarReminders.Calendar = mvarMe.CalendarControl
         g_dlgCalendarReminders.ParentHWND = mvarMe.hwnd
         g_dlgCalendarReminders.RemindersWindowShowInTaskBar = False
    
         g_dlgCalendarReminders.CreateRemindersWindow
      End If
   End With
      
   Screen.MousePointer = GetTag(mvarMe, "MousePointer", Screen.MousePointer)
End Sub
Private Sub PageSetup()
    mvarMe.CalendarControl.ShowPrintPageSetup
End Sub
Private Sub PrintCalendar()
   With mvarMe.CalendarControl
      ' Header ---------------------------------------------------
      
      '.PrintOptions.Header.Font.Name = "Courier New"
      '.PrintOptions.Header.Font.Bold = True
      '.PrintOptions.Header.Font.SIZE = 10
        
      .PrintOptions.Header.TextLeft = "(c)1998-2007 Codejock Software, All Rights Reserved."
      .PrintOptions.Header.TextCenter = "Calendar Control"
      .PrintOptions.Header.TextRight = "Page 1 of 1 "
    
   ' Footer ---------------------------------------------------
    
   '.PrintOptions.Header.Font.Italic = True
    
      .PrintOptions.Footer.TextLeft = "Date: " & DateValue(Now)
      .PrintOptions.Footer.TextCenter = "Codejock Software, " & vbLf & " Print calendar example "
      .PrintOptions.Footer.TextRight = "Time: " & TimeValue(Now)
    
   ' Other fonts
    
   '.PrintOptions.DateHeaderFont.SIZE = 12
   '.PrintOptions.DateHeaderCalendarFont.SIZE = 10
            
   'If .ShowPrintPageSetup() Then
      .PrintCalendar 0
   'End If
   End With
End Sub
Private Sub PrintPreview()
   With mvarMe
      .CalendarControl.PrintPreviewOptions.Title = "Calendário D'pil"
      .CalendarControl.PrintPreview True
   End With
End Sub
Private Sub ExibirProducao()
   Dim Sql As String
   Dim sMsg As String
   Dim nPreco As Double
   Dim nVLVENDA As Double
   Dim nQtdVenda As Double
   Dim nQtdVendaPrev As Double
   
   If Not AcessopEspecial(mvarSys, "UTIL_VERPROD") Then
      Exit Sub
   End If
   Sql = "Select IDPROD, VLVENDA "
   Sql = Sql & " From SPRODUTO "
   Sql = Sql & " Where ESERVICO=1 "
   Sql = Sql & " And IDLOJA=" & mvarIDLOJA
   Sql = Sql & " Order By IDPROD"
   
   Sql = "Select (Sum(I.QTDVENDA*I.VLUNIT)-Sum(V.VLDESC))/Sum(I.QTDVENDA) as [VLVENDA]"
   Sql = Sql & "  From CITENSVENDA I"
   Sql = Sql & " Join CVENDA V On V.IDLOJA=I.IDLOJA And V.IDVENDA=I.IDVENDA"
   Sql = Sql & " Join OCLIENTE C On C.IDLOJA=V.IDLOJA And C.IDCLIENTE=V.IDCLIENTE And C.ISENTO=0"
   Sql = Sql & " Join SPRODUTO P On P.IDLOJA=I.IDLOJA And P.IDPROD=I.IDPROD"
   Sql = Sql & " Where V.DTVENDA >= GetDate() - 30"
   Sql = Sql & " And V.IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And P.ESERVICO=1"
   Sql = Sql & " And P.EVENDA=1"
   Sql = Sql & " Group by P.ESERVICO"
   If mvarSys.xDb.Abretabela(Sql) Then
      nPreco = xVal(mvarSys.xDb.RsAux("VLVENDA"))
   End If
   
   Sql = "Select P.ESERVICO, SUM(I.QTDVENDA) [QTD], SUM(I.QTDVENDA*I.VLUNIT) [VALOR]" & vbNewLine
   Sql = Sql & " From CITENSVENDA I" & vbNewLine
   Sql = Sql & " Join CVENDA V On V.IDLOJA=I.IDLOJA And V.IDVENDA=I.IDVENDA" & vbNewLine
   Sql = Sql & " Join SPRODUTO P On I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD AND P.EVENDA=1" & vbNewLine
   Sql = Sql & " Where I.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And Month(V.DTVENDA)=" & SqlNum(Month(Now())) & vbNewLine
   Sql = Sql & " And Year(V.DTVENDA)=" & SqlNum(Year(Now())) & vbNewLine
   Sql = Sql & " And P.ESERVICO=1"
   Sql = Sql & " Group By P.ESERVICO"
   If mvarSys.xDb.Abretabela(Sql) Then
      nVLVENDA = xVal(mvarSys.xDb.RsAux("VALOR"))
   End If
   
   sMsg = ""
   sMsg = sMsg & "Previsão Inicial : " & GetTag(mvarPane.PctPrev0, "VALOR", 0) & Space(5) & vbNewLine
   sMsg = sMsg & "===========================" & vbNewLine
   sMsg = sMsg & "========= PRODUÇÃO =========" & vbNewLine
   sMsg = sMsg & "===========================" & vbNewLine
   sMsg = sMsg & "Realizado          : " & GetTag(mvarPane.PctReal, "VALORPROD", 0) & Space(5) & " / R$ " & ValBr(GetTag(mvarPane.PctReal, "VALORPROD", 0) * nPreco) & vbNewLine
   sMsg = sMsg & "Previsão Atual  : " & GetTag(mvarPane.PctPrev, "VALORPROD", 0) & Space(5) & " / R$ " & ValBr(GetTag(mvarPane.PctPrev, "VALORPROD", 0) * nPreco) & vbNewLine & vbNewLine
   
   sMsg = sMsg & "===========================" & vbNewLine
   sMsg = sMsg & "=========== VENDA ==========" & vbNewLine
   sMsg = sMsg & "===========================" & vbNewLine
   nQtdVenda = GetTag(mvarPane.PctReal, "VALORVENDA", 0)
   nQtdVendaPrev = GetTag(mvarPane.PctPrev, "VALORVENDA", 0)
   sMsg = sMsg & "Realizado          : " & nQtdVenda & Space(5) & " / R$ " & ValBr(nQtdVenda * nPreco) & vbNewLine
   sMsg = sMsg & "Previsão Atual  : " & nQtdVendaPrev & Space(5) & " / R$ " & ValBr(nQtdVendaPrev * nPreco) & vbNewLine
   
   
   Call ExibirInformacao(sMsg, "Produção")
End Sub
Public Sub ExibeAgendas()
   Dim pData      As CalendarDataProvider
   Dim pSchedules As CalendarSchedules
   Dim pRCDesc    As CalendarResourceDescription
   Dim bResult    As Boolean
   Dim i          As Integer
   Dim vLoja      As Variant
   Dim vSala      As Variant
   Dim ConnString As String
   'Dim sSalas     As String
   Dim bAux       As Boolean
   Dim sAux       As String
   
   bResult = True
   ConnString = "Provider=custom;DSN=CalendarSQL_Server"
   
   bAux = True
   If g_DataResourcesMan Is Nothing Or bAux Then
      bAux = True
      Set g_DataResourcesMan = New CalendarResourcesManager
      bResult = g_DataResourcesMan.AddDataProvider(ConnString, xtpCalendarDPF_CreateIfNotExists + xtpCalendarDPF_SaveOnDestroy + xtpCalendarDPF_CloseOnDestroy)
   End If
   
   If SalasAtivas <= 0 Then Exit Sub
   If Not bResult Then Exit Sub
   
   
   Set pData = g_DataResourcesMan.DataProvider(0)
   Set pSchedules = pData.Schedules
   If pSchedules Is Nothing Then Exit Sub
   
   bAux = True
   If bAux Then
      For i = pSchedules.Count To 1 Step -1
         pSchedules.RemoveSchedule i
      Next
      
      If pSchedules.Count < 1 Then
         For Each vLoja In gUnidades
            For Each vSala In vLoja.Salas
               pSchedules.AddNewSchedule vSala.NMSALA
               pSchedules(pSchedules.Count - 1).Id = vSala.ScheduleID
            Next
         Next
         pData.Save
      End If
      If pSchedules.Count < 1 Then
         For Each vLoja In gUnidades
            For Each vSala In vLoja.Salas
               pSchedules.AddNewSchedule vSala.NMSALA
               pSchedules(pSchedules.Count - 1).Id = vSala.ScheduleID
               If vSala.IDLOJA = mvarIDLOJA Then
                  pSchedules(pSchedules.Count - 1).Id = vSala.IDSALA
               End If
            Next
         Next
         pData.Save
      End If
   End If
   For i = g_DataResourcesMan.ResourcesCount - 1 To 0 Step -1
      g_DataResourcesMan.RemoveResource i
   Next
   If g_DataResourcesMan.ResourcesCount = 0 Then
      If Not ExisteItem(gUnidades, CStr(mvarIDLOJA)) Then
         sAux = IIf(mvarSys.GetParam("EXIBESALA", Default:=1) = 0, "01", "Sala 01")
         If g_DataResourcesMan.AddResource(sAux, True) Then
            Set pRCDesc = g_DataResourcesMan.Resource(g_DataResourcesMan.ResourcesCount - 1)
            pRCDesc.Resource.SetDataProvider pData, True
            pRCDesc.Resource.ScheduleIDs.Add mvarIDLOJA * 1000 + 1
            pRCDesc.GenerateName = True
         End If
      Else
         For Each vSala In gUnidades(CStr(mvarIDLOJA)).Salas 'vLoja.Salas
            If vSala.Selecionada Then
               If g_DataResourcesMan.AddResource(vSala.NMSALA, True) Then
                  Set pRCDesc = g_DataResourcesMan.Resource(g_DataResourcesMan.ResourcesCount - 1)
                  pRCDesc.Resource.SetDataProvider pData, True
                  pRCDesc.Resource.ScheduleIDs.Add vSala.ScheduleID
                  pRCDesc.GenerateName = True
               End If
            End If
         Next
      End If
      '// Save changed resources configuration
      g_DataResourcesMan.SaveCfg ConnString
   End If
    '// Apply resources configuration
   If bAux Then
      g_DataResourcesMan.ApplyToCalendar mvarMe.CalendarControl
   End If
End Sub
Private Function SalasAtivas() As Integer
   Dim sSalas As String
   Dim vSala  As Variant
   'SalasAtivas = gUnidades(CStr(mvarIDLOJA)).Salas.Count
   If Not ExisteItem(gUnidades, CStr(mvarIDLOJA)) Then
     sSalas = "1"
     SalasAtivas = 1
      mvarSys.Propriedades("SALAS") = sSalas
   Else
      If gUnidades(CStr(mvarIDLOJA)).Salas.Count > 0 Then
         sSalas = ""
         mvarSys.Propriedades("SALAS") = ""
         For Each vSala In gUnidades(CStr(mvarIDLOJA)).Salas
            sSalas = sSalas & IIf(Trim(sSalas) = "", "", ", ") & vSala.IDSALA
         Next
         sSalas = IIf(Trim(sSalas) = "", "0", sSalas)
         mvarSys.Propriedades("SALAS") = sSalas
      End If
      SalasAtivas = gUnidades(CStr(mvarIDLOJA)).Salas.Count
   End If
   
End Function
Private Sub CreateTaskPanel()
   Dim Group As TaskPanelGroup
   Dim Item As TaskPanelGroupItem
   
   
   mvarPane.FraCliente.Height = 1500
   mvarPane.FraResumoDia.Height = mvarPane.FraCliente.Height
   
   
   With mvarPane.wndTaskPanel
      .Behaviour = xtpTaskPanelBehaviourList
      '.SingleSelection = True
      '.HotTrackStyle = xtpTaskPanelHighlightText
   
      '.SetImageList mvarPane.imlTaskPanelIcons
      .ColorSet.BackgroundLight = mvarPane.BackColor
      .ColorSet.BackgroundDark = mvarPane.BackColor
      .Expandable = True
      
'      Set Group = .Groups.Add(0, "Calendário")
'      Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
'       Set Item.Control = mvarPane.wndDatePicker
'       'mvarPane.FraFiltro.BackColor = Item.BackColor
'      Group.Expandable = True
'      Group.Expanded = True
      
      
      Set Group = .Groups.Add(0, "Filtros")
      Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
       Set Item.Control = mvarPane.FraFiltro
       mvarPane.FraFiltro.BackColor = Item.BackColor
      Group.Expandable = True
      Group.Expanded = True
      
'      Set Group = .Groups.Add(0, "Ações de Agenda")
'      Group.Items.Add 1, "Criar/Abrir Compromisso", xtpTaskItemTypeLink, 17
'      Group.Items.Add 2, "Criar/Abrir Lista de Espera", xtpTaskItemTypeLink, 18
'      Group.Items.Add 3, "Excluir Item", xtpTaskItemTypeLink, 19
'      Group.Expandable = True
'      Group.Expanded = False
       
'      Set Group = .Groups.Add(0, "Dados de Cliente")
'      Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
'       Set Item.Control = mvarPane.FraCliente
'       mvarPane.FraCliente.BackColor = Item.BackColor
'      Group.Expandable = True
'      Group.Expanded = False
      
       
'      Set Group = .Groups.Add(0, "Resumo do Dia")
'      Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
'      Set Item.Control = mvarPane.FraResumoDia
'      mvarPane.FraResumoDia.BackColor = Item.BackColor
'      Group.Expandable = True
'      Group.Expanded = False
    
      'Set Group = mvarPane.wndTaskPanel.Groups.Add(0, "Ferramentas")
      'Set Item = Group.Items.Add(ID_TASKITEM_MALADIRETA, "Mala Direta", xtpTaskItemTypeLink)
      'Group.Expanded = False
      
      .Groups(1).Expanded = True
      .Behaviour = xtpTaskPanelBehaviourList
      .SingleSelection = True
      .Reposition
      
      .Visible = True
   End With
   
End Sub

Private Sub mvarPane_wndDatePickerDayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
   If Weekday(Day) = vbSunday Then
        Metrics.ForeColor = vbRed
   End If
   If Format(Day, "dd/mm/yy") = Format(Now() + 28, "dd/mm/yy") Then
      Metrics.ForeColor = &H8000&
      Metrics.Font.Underline = True
   End If
   'If Format(Day, "dd/mm/yy") = Format(Now() + 21, "dd/mm/yy") Then
   '   'Metrics.ForeColor = &H8000&
   '   Metrics.Font.Italic = True
   '   Metrics.Font.Underline = True
   'End If
End Sub

Private Sub mvarPane_WndTaskPanelItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
   Dim i As Integer
   With mvarPane.wndTaskPanel
      For i = 0 To .Groups.Count
         If i <> Item.Id Then
            .Groups(i).Expanded = False
         End If
      Next
   End With
End Sub
