VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Tarefas"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const StrPesquisa = "Pesquisar Tarefa"
Private WithEvents mvarMe     As FrmTarefas
Attribute mvarMe.VB_VarHelpID = -1
Private WithEvents mvarCmmdBars  As XtremeCommandBars.CommandBars
Attribute mvarCmmdBars.VB_VarHelpID = -1
Private mvarTLPane As TL_PaneTarefa
Attribute mvarTLPane.VB_VarHelpID = -1

Private Enum eMenuCal
   Mnu_Tarefas = 2300
   Mnu_TarEditar
   Mnu_TarSit
   Mnu_TarMail
 
    'Mnu_TarefaEditar
   Mnu_TarEditarEntrar
   Mnu_TarEditarNova
   Mnu_TarEditarExcluir
   
   'Mnu_TarefaSituacao
   Mnu_TarSitNormal
   Mnu_TarSitAndamento
   Mnu_TarSitConcluido
End Enum
Private Enum ePopTar
   Pop_TarEdit = 300
   Pop_TarNew
   Pop_TarSit
End Enum


Dim iCampo     As Collection
Dim fntStrike  As StdFont
'Private bNovo As Boolean

Private RsLan        As Object
Private mvarSys      As Object
Private mvarIDLOJA   As Integer
Private mvarIDTAREFA As Long
Private mvarTLCal    As Object
Public Property Set TLPane(ByVal vData As TL_PaneTarefa)
   Set mvarTLPane = vData
End Property
Public Property Set TLCal(ByVal vData As Object)
   Set mvarTLCal = vData
End Property
Public Property Get TLCal() As Object
   Set TLCal = mvarTLCal
End Property
Public Property Get frm() As Object
   Set frm = mvarMe
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   Set gSys = vData
   mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
   Set mvarCmmdBars = mvarSys.MDI.CommandBars
End Property
Public Property Get Sys() As Object
   Set Sys = mvarSys
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDTAREFA(ByVal vData As Long)
   mvarIDTAREFA = vData
End Property
Public Property Get IDTAREFA() As Long
   IDTAREFA = mvarIDTAREFA
End Property
Public Sub Show(Optional Reload As Boolean = False)
   If Reload Then
      Set Sys = Sys
'      Unload mvarMe
   End If
   
   '* Carregar e Acoplar Form no MDI
   Load mvarMe
   Call AcoplarForm(mvarMe, 2, mvarSys)
   mvarMe.Show

   '* Exibir Menu no MDI
   If Not mvarSys.MDI.CommandBars.FindControl(, Mnu_Tarefas, False) Is Nothing Then
      mvarSys.MDI.CommandBars.FindControl(, Mnu_Tarefas).Visible = True
   End If

   '* Definir foco
   With mvarMe.GrdTarefas
      DoEvents
      If .Visible And .Enabled Then
         'Call MontaTela
         Call PopulaGrdTarefas
         .SetFocus
      End If
   End With
   If Reload Then Call mvarMe_Rezise
End Sub
Public Sub Hide()
   mvarMe.Hide
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmTarefas
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub

Private Sub mvarCmmdBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   Dim n As Object
   
   DoEvents
   Screen.MousePointer = vbHourglass
   Select Case UCase(Control.Category)
      Case "TAREFA"
         Select Case Control.Id
            Case Mnu_TarSitNormal, Mnu_TarSitAndamento, Mnu_TarSitConcluido
               Call AlterarSitTarefa(Control.Id - Mnu_TarSitNormal)
               
            Case Mnu_TarEditarNova
               
            Case Mnu_TarMail
               Call ExibireMailPadrao
               
         End Select
   End Select
   Screen.MousePointer = vbDefault
End Sub
Private Sub AlterarSitTarefa(pIndex As Long)
   Dim TbTarefa As Object
   Dim nIdTarefa As Long
   Dim i As Integer
   Dim xItem As ReportRecordItem

   nIdTarefa = xVal(GetTag(mvarMe.GrdTarefas, "IDTAREFA", 0))
   If nIdTarefa <> 0 Then
      Set TbTarefa = CriarObjeto("BANCO_3R.TB_OTAREFAEVT", False)
      Set TbTarefa.xDb = mvarSys.xDb
      If TbTarefa.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDTAREFA:=nIdTarefa) Then
         Select Case pIndex
            Case 0: TbTarefa.SITTAREFA = 0 '"Não Iniciada", 0
            Case 1: TbTarefa.SITTAREFA = 5 '"Em Andamento", 5
            Case 2: TbTarefa.SITTAREFA = 10 '"Concluída", 10
         End Select
         If TbTarefa.IsDirt Then
            If TbTarefa.Salvar Then
               Dim oCAD As TL_CadTarefa
               Set oCAD = New TL_CadTarefa
               Set oCAD.Sys = mvarSys
               oCAD.IDLOJA = mvarIDLOJA
               oCAD.IDTAREFA = nIdTarefa
               Call oCAD.ConfirmaAgenda(mvarSys, mvarIDLOJA, nIdTarefa, mvarTLCal)

               Call mvarMe_KeyUp(vbKeyF5, 0)
            
               On Error Resume Next
'               mvarMe.GrdTarefas.FocusedRow.Selected = True
'               'mvarMe.GrdTarefas.SelectedRows(0).Record
'               i = mvarMe.GrdTarefas.SelectedRows.Count - 1
'               Set xItem = mvarMe.GrdTarefas.SelectedRows(i).Record.Item(0)
'               mvarMe.Refresh
'               If Not xItem Is Nothing Then
'                  If TbTarefa.SITTAREFA = 0 Then
'                     xItem.Icon = -1
'                  ElseIf TbTarefa.SITTAREFA = 5 Then
'                     xItem.Icon = 3
'                  ElseIf TbTarefa.SITTAREFA = 10 Then
'                     xItem.Icon = 1
'                  End If
'                  mvarMe.Refresh
'                  mvarMe.GrdTarefas.Redraw
'               End If
            End If
         End If
      End If
   End If
End Sub
Private Sub ExibireMailPadrao()
   Dim oTarefa As TL_CadTarefa
   Dim nIdTarefa As Long

   nIdTarefa = xVal(GetTag(mvarMe.GrdTarefas, "IDTAREFA", 0))
   If nIdTarefa <> 0 Then
      Set oTarefa = New TL_CadTarefa
      Set oTarefa.Sys = mvarSys
      If oTarefa.ExibireMail(nIdTarefa) Then
         Dim TbTarefa As Object
         Set TbTarefa = CriarObjeto("BANCO_3R.TB_OTAREFAEVT")
         With TbTarefa
            Set .xDb = mvarSys.xDb
            If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDTAREFA:=nIdTarefa) Then
               .SITTAREFA = 10 '* Concluída
               .DSCTAREFA = .DSCTAREFA & vbNewLine & "[" & Format(mvarSys.xDb.SysDate(3), "dd/mm/yy hh:mm") & "] - e-Mail enviado. "
               If .IsDirt Then
                  If .Salvar Then
                     Call mvarMe_KeyUp(vbKeyF5, 0)
                  End If
               End If
            End If
         End With
         Set TbTarefa = Nothing
      End If
   End If
End Sub
Private Sub mvarMe_Activate()
'   Call MontaTela
'   Call PopulaTela
   If GetTag(mvarMe, "LOAD", 1) = 1 Then
      Call SetTag(mvarMe, "LOAD", 0)
      Call mvarMe_GrdTarefasSelectionChanged
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmdEMAILClick()
   Dim oReg    As ReportRecord
   Dim sTipo   As String
   Dim nTipo   As Boolean
      
   nTipo = 0
   With mvarMe
      If .GrdTarefas.SelectedRows.Count = 1 Then
         Set oReg = .GrdTarefas.SelectedRows(0).Record
         If Not oReg Is Nothing Then
            sTipo = UCase(Trim(oReg(iCampo("TIPO")).Value))
            If InStr(sTipo, "-") > 0 Then
               sTipo = Trim(Mid(sTipo, 1, InStr(sTipo, "-") - 1))
               nTipo = xVal(sTipo)
            End If
         End If
      End If
   End With
   Call ExibireMail(nTipo, mvarMe.CmdEMAIL.Caption)
End Sub
Private Sub ExibireMail(nTipo, sMsgTo)
   If Between(nTipo, eTpTarefa.TarBoasVindas, eTpTarefa.TarNaoMarcado) Then
      Call ExecuteLink("mailto:" & sMsgTo)
   Else
      Call ExecuteLink("mailto:" & sMsgTo)
   End If
End Sub
Private Sub mvarMe_GrdTarefasBeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
   
   If Item Is Nothing Then Exit Sub
'   If Not iCampo Is Nothing Then
'      If Item.Index = iCampo("!") Then
'         If Item.Value = "0" Then Item.Value = ""
'      End If
'
'      If Item.Index = iCampo("Sit.") Then
         If Row.Record(iCampo("FLGDELETE")).Value = 1 Or Row.Record(iCampo("SITTAREFA")).Value = 10 Then 'And (Item.Caption <> " ") Then
'            Item.Icon = 2
'            Metrics.ForeColor = vbRed
'            Metrics.Font.Strikethrough = True
            Set Metrics.Font = fntStrike
         Else
'            If Row.Record(iCampo("SITTAREFA")).Value = 0 Then
'               Metrics.ForeColor = 0
'               Metrics.Font.Strikethrough = False
'            Else
'               Metrics.ForeColor = RGB(128, 128, 128)
'               Metrics.Font.Strikethrough = True
'            End If
            'If Row.Record(iCampo("SITTAREFA")).Value = 0 Then
            '   Item.Icon = Nothing
            'ElseIf Row.Record(iCampo("SITTAREFA")).Value = 5 Then
            '   Item.Icon = 3
            'ElseIf Row.Record(iCampo("SITTAREFA")).Value = 10 Then
            '   Item.Icon = 1
            'End If
         End If
'      End If
'
'   End If
   If ((Row.Index Mod 2 = 0) And (Row.GroupRow = False)) Then
       Metrics.BackColor = RGB(245, 245, 245)
   Else
      'Metrics.BackColor = RGB(65, 165, 241)
   End If
   
End Sub
Private Sub mvarMe_GrdTarefasMouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
   Dim Popup     As CommandBar
   Dim xItem As ReportRecordItem
   Dim xRow As ReportRow
   
   Call SetTag(mvarMe.GrdTarefas, "IDTAREFA", 0)
   mvarMe.GrdTarefas.SelectedRows.DeleteAll
   If Button = 2 Then
      Set xItem = mvarMe.GrdTarefas.HitTest(x, y).Item
      Set xRow = mvarMe.GrdTarefas.HitTest(x, y).Row
      If Not xItem Is Nothing Then
         DoEvents
         xRow.Selected = True
         Call PopulaTab
         Call SetTagGrid(xItem, xRow)
         DoEvents
         Set Popup = mvarSys.MDI.CommandBars.ContextMenus.Find(Pop_TarSit)
         If Not Popup Is Nothing Then
            Popup.Controls(4).Enabled = Between(xVal(GetTag(mvarMe.GrdTarefas, "IDTPTAREFA")), eTpTarefa.TarBoasVindas, eTpTarefa.TarNaoMarcado)
            Popup.Controls(4).Enabled = Popup.Controls(4).Enabled And (Trim(GetTag(mvarMe.GrdTarefas, "EMAIL")) <> "")
            Popup.ShowPopup
         End If
      End If
   End If
End Sub
Private Sub SetTagGrid(Optional xItem As ReportRecordItem, Optional xRow As ReportRow)
Dim nAux As Long
   If xRow Is Nothing Then
      If mvarMe.GrdTarefas.SelectedRows.Count > 0 Then
         Set xRow = mvarMe.GrdTarefas.SelectedRows(0)
      End If
   End If
   If xItem Is Nothing And Not xRow Is Nothing Then
      If Not xRow.GroupRow Then
         Set xItem = xRow.Record(0)
      End If
   End If
   If Not xItem Is Nothing Then
      Dim i As Integer
      For i = 0 To iCampo.Count - 1
         Call SetTag(mvarMe.GrdTarefas, mvarMe.GrdTarefas.Columns(i).Caption, xItem.Record(i).Value)
      Next
      Call SetTag(mvarMe.GrdTarefas, "IDTAREFA", xItem.Record(iCampo("IDTAREFA")).Value)
      Call SetTag(mvarMe.GrdTarefas, "IDTPTAREFA", xVal(Mid(xItem.Record(iCampo("Tipo")).Value, 1, 3)))
      Call SetTag(mvarMe.GrdTarefas, "ROW", xRow.Index)
   End If
End Sub
Private Sub mvarMe_GrdTarefasRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   mvarMe.GrdTarefas.SelectedRows.DeleteAll
   Row.Selected = True
   Call EditarObjeto
End Sub
Private Sub mvarMe_GrdTarefasSelectionChanged()
   Call LimpaTab
   Call PopulaTab
   Call SetTagGrid
   DoEvents
   mvarMe.GrdTarefas.SetFocus
End Sub
Private Sub LimpaTab()
   With mvarMe
      .LblNMCLIENTE.Caption = ""
      .LblTEL1.Caption = ""
      .LblTEL2.Caption = ""
      .LblFAX.Caption = ""
      .LblEMAIL.Caption = ""
      .CmdEMAIL.Caption = ""
      .LblSTARTTIME.Caption = ""
      .LblEvento.Caption = "Evento: "
      .LblSituacao.Caption = ""
      .TxtOutros.Text = ""
   End With
End Sub
Private Sub PopulaTab()
   Dim nPag As Currency
   Dim nDep As Currency
   Dim oReg As ReportRecord
   Dim sTel1 As String
   Dim sTel2 As String
   Dim sFax  As String
   Dim sTipo As String
   Dim sAux  As String

   With mvarMe
      If .GrdTarefas.SelectedRows.Count <> 1 Then Exit Sub

      Set oReg = .GrdTarefas.SelectedRows(.GrdTarefas.SelectedRows.Count - 1).Record
      If oReg Is Nothing Then Exit Sub
      
      sTel1 = UnFormat(Trim(oReg(iCampo("TEL1")).Value))
      sTel1 = Mid(sTel1, 1, 4) & "-" & Mid(sTel1, 5)
      sTel2 = UnFormat(Trim(oReg(iCampo("TEL2")).Value))
      sTel2 = Mid(sTel2, 1, 4) & "-" & Mid(sTel2, 5)
      sFax = UnFormat(Trim(oReg(iCampo("FAX")).Value))
      sFax = Mid(sFax, 1, 4) & "-" & Mid(sFax, 5)
      
      sTipo = UCase(Trim(oReg(iCampo("TIPO")).Value))
      If InStr(sTipo, "-") > 0 Then
         sTipo = Trim(Mid(sTipo, InStr(sTipo, "-") + 1))
      End If
      mvarMe.LblDSCTPTAREFA.Caption = sTipo
      
      mvarMe.LblNMCLIENTE.Caption = oReg(iCampo("NOME")).Value
      mvarMe.LblTEL1.Caption = sTel1
      mvarMe.LblTEL2.Caption = sTel2
      mvarMe.LblFAX.Caption = sFax
      mvarMe.LblEMAIL.Caption = oReg(iCampo("EMAIL")).Value
      mvarMe.CmdEMAIL.Caption = mvarMe.LblEMAIL.Caption
      mvarMe.CmdEMAIL.Width = mvarMe.LblEMAIL.Width
      
      mvarMe.LblSTARTTIME.Caption = Format(oReg(iCampo("STARTDATETIME")).Value, "dd/mm/yy hh:mm") & "h"
      If IsDate(oReg(iCampo("STARTDATETIME")).Value) Then
         sAux = "Sessão: "
         Select Case Weekday(CDate(oReg(iCampo("STARTDATETIME")).Value))
            Case 1: sAux = sAux & "Domingo"
            Case 2: sAux = sAux & "Segunda-Feira"
            Case 3: sAux = sAux & "Terça-Feira"
            Case 4: sAux = sAux & "Quarta-Feira"
            Case 5: sAux = sAux & "Quinta-Feira"
            Case 6: sAux = sAux & "Sexta-Feira"
            Case 7: sAux = sAux & "Sábado"
         End Select
         mvarMe.LblSTARTTIME.AutoSize = True
         mvarMe.LblSTARTTIME.Left = mvarMe.GrpEVENTO.Width - mvarMe.LblSTARTTIME.Width - 300
         mvarMe.LblEvento.Caption = sAux & ", "
         mvarMe.LblEvento.Left = mvarMe.LblSTARTTIME.Left - mvarMe.LblEvento.Width
      End If
      
      If oReg(iCampo("FLGCANCELADO")).Value = 1 Then
         mvarMe.LblSituacao.Caption = "Cancelado"
         mvarMe.LblSituacao.ForeColor = &HC0C0FF
      ElseIf oReg(iCampo("FLGCONFIRMADO")).Value = 1 Then
         mvarMe.LblSituacao.Caption = "Confirmado"
         mvarMe.LblSituacao.ForeColor = RGB(140, 180, 120) '&HE0E0E0
      End If
      
      mvarMe.TxtOutros.Width = mvarMe.LblEvento.Left - mvarMe.TxtOutros.Left - 240
      mvarMe.TxtOutros.Text = oReg(iCampo("DSCTAREFA")).Value
      'mvarMe.TxtOutros.SelStart = Len(mvarMe.TxtOutros.Text)
      'mvarMe.TxtOutros.SelLength = 0
      'Call ScrollText(mvarMe.TxtOutros.hwnd, 7)
      
      'E.FLGREMARCADO
   End With
End Sub
Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim xIDTarefa As Long
   Dim iRow As Long
   Dim xRow As ReportRow
   
   DoEvents
   If KeyCode = vbKeyF5 Then
      If mvarMe.GrdTarefas.SelectedRows.Count > 0 Then
         If Not mvarMe.GrdTarefas.SelectedRows(0).Record Is Nothing Then
            xIDTarefa = mvarMe.GrdTarefas.SelectedRows(0).Record.Item(iCampo("IDTAREFA")).Value
         End If
      End If
      Call PopulaGrdTarefas
      
      If xIDTarefa > 0 Then
         For iRow = 0 To mvarMe.GrdTarefas.Records.Count - 1
            If mvarMe.GrdTarefas.Records(iRow).Item(iCampo("IDTAREFA")).Value = xIDTarefa Then
               mvarMe.GrdTarefas.Records(iRow).Expanded = True
               Set xRow = mvarMe.GrdTarefas.Rows.FindRowInTree(mvarMe.GrdTarefas.Records(iRow))
               Set xRow = xRow.ParentRow
               While Not xRow Is Nothing
                  xRow.Expanded = True
                  Set xRow = xRow.ParentRow
               Wend
               Set xRow = mvarMe.GrdTarefas.Rows.FindRowInTree(mvarMe.GrdTarefas.Records(iRow))
               xRow.EnsureVisible
               xRow.Selected = True
               mvarMe.GrdTarefas.Redraw
               Exit For
            End If
         Next
      End If
   ElseIf KeyCode = vbKeyDelete Then
      If mvarMe.ActiveControl Is mvarMe.GrdTarefas Then
         'eXCLUI
      End If

   ElseIf KeyCode = vbKeyEscape Then
      If mvarMe.txtFiltrar.Text <> StrPesquisa Then
         mvarMe.txtFiltrar.Text = StrPesquisa
         Call PopulaGrdTarefas
      End If
      Call PopulaTab

   ElseIf KeyCode = vbKeyInsert Then
      'nOVO

   ElseIf (Asc(UCase(Chr(KeyCode))) >= Asc("A") And Asc(UCase(Chr(KeyCode))) >= Asc("Z")) Or (Asc(UCase(Chr(KeyCode))) >= Asc("0") And Asc(UCase(Chr(KeyCode))) >= Asc("9")) Then
      'If Not mvarMe.ActiveControl Is mvarMe.txtFiltrar Then
      '   mvarMe.txtFiltrar.Text = UCase(Chr(KeyCode))
      '   mvarMe.txtFiltrar.SetFocus
      '   mvarMe.txtFiltrar.SelStart = Len(mvarMe.txtFiltrar.Text)
      'End If
   End If
End Sub
Private Sub mvarMe_LblEMAILClick()
   Call ExecuteLink("mailto:" & mvarMe.LblEMAIL.Caption)
End Sub
Private Sub mvarMe_Load()
   
   Call MontarTela
   Call MontarMenu
   'Call MontarToolbar
   'Call MontarStatusBar
   
   
   Call PopulaTela
   DoEvents
End Sub
Public Sub PopulaTela()
   Call LimpaGrd
   Call PopulaGrdTarefas
End Sub
Private Sub MontarTela()
   'ReportControlGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   With mvarMe
      .GrpEVENTO.ZOrder 1

      .txtFiltrar = StrPesquisa
      .SccConta.VisualTheme = xtpShortcutThemeOffice2003
      .SccConta.GradientHorizontal = False
      With .GrdTarefas
         .Icons = mvarMe.ImgShortcutBar.Icons
         Set fntStrike = .PaintManager.TextFont
         fntStrike.Strikethrough = True
         
         .BorderStyle = xtpBorderFrame
         .PaintManager.ColumnStyle = xtpColumnOffice2007
         .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
         .PaintManager.ShadeGroupHeadings = False
         .PaintManager.GroupRowTextBold = True
         '.PaintManager.GroupBoxBackColor =vbred
         .PaintManager.GroupForeColor = &H4000&
         .PaintManager.GroupShadeBackColor = mvarMe.BackColor
         '.PaintManager.GroupShadeForeColor = vbRed
         '.PaintManager.GroupShadeBorderColor = vbBlue
         .PaintManager.RefreshMetrics

         .FooterRowsAllowAccess = True
         .FooterRowsAllowEdit = True
         .FooterRowsEnableSelection = True

         '.SetImageList mvarMe.ImageList1
         .SetCustomDraw xtpCustomBeforeDrawRow

         .MultipleSelection = True
         .AutoColumnSizing = False
         .AllowColumnRemove = False
         .AllowColumnReorder = True
         .AllowColumnResize = True
         .AllowColumnSort = True
         .AllowEdit = False
         
   
         .ShowFooter = False
         .ShowFooterRows = False
         .ShowGroupBox = False
         .ShowHeader = True
         .ShowHeaderRows = True
         .ShowItemsInGroups = False
         .ShowRowFocus = True
      End With
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
         Set oCommBar = .Add("Menu Tarefas", xtpBarTop)
         oCommBar.Closeable = False
         oCommBar.Customizable = False
         oCommBar.EnableDocking xtpFlagHideWrap
      Else
         Set oCommBar = .ActiveMenuBar
      End If
      
      Set oMenuMain = oCommBar.FindControl(, Mnu_Tarefas)
      If Not oMenuMain Is Nothing Then
         If (oMenuMain.Category = "TAREFA") Then
            oMenuMain.Visible = True
            GoTo Saida
         End If
      End If
      
      With oCommBar
         '************
         '* Menu Tarefa
         Set oMenuMain = AddButtonBar(.Controls, Mnu_Tarefas, "Tarefas", False, xtpControlPopup, "")
         With oMenuMain.CommandBar
            AddButtonBar .Controls, Mnu_TarEditarNova, "Nova"
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
Private Sub MontarContextMenu()
   Dim ContextMenu   As CommandBar
   Dim oMenuItem     As CommandBarControl
   
   '**************
   '* Pop-up Edit
   Set ContextMenu = mvarSys.MDI.CommandBars.ContextMenus.Add(Pop_TarEdit, "Context Menu Edit")
   With ContextMenu
      AddButtonBar .Controls, Mnu_TarEditarNova, "Nova", Category:="TAREFA"
      AddButtonBar .Controls, Mnu_TarEditarEntrar, "Abrir", Category:="TAREFA"
      AddButtonBar .Controls, Mnu_TarEditarExcluir, "Excluir", Category:="TAREFA"
   End With
   ContextMenu.Controls.CreateOriginalControls
   
      
   '**************
   '* Pop-up Situação
   Set ContextMenu = mvarSys.MDI.CommandBars.ContextMenus.Add(Pop_TarSit, "Context Menu Status")
   With ContextMenu
'      Set oMenuItem = AddButtonTar(.Controls, Mnu_TarSit, "Situação", False, xtpControlButtonPopup)
'      With oMenuItem
         AddButtonBar .Controls, Mnu_TarSitNormal, "Não Iniciada", False, xtpControlButton, Category:="TAREFA"
         AddButtonBar .Controls, Mnu_TarSitAndamento, "Andamento", False, xtpControlButton, Category:="TAREFA"
         AddButtonBar .Controls, Mnu_TarSitConcluido, "Concluído", False, xtpControlButton, Category:="TAREFA"
'      End With
      AddButtonBar .Controls, Mnu_TarMail, "Enviar e-Mail padrão", True, xtpControlButton, Category:="TAREFA"
      
   End With
   ContextMenu.Controls.CreateOriginalControls
   
End Sub
Private Sub mvarMe_Rezise()
   On Error Resume Next

   mvarMe.GrpEVENTO.ZOrder 1
'   mvarMe.Refresh
   mvarMe.SccConta.Move 0, 50, mvarMe.ScaleWidth - 50, 420
   With mvarMe.txtFiltrar
      .Left = mvarMe.Width - 2800
      .Top = 90
   End With
   With mvarMe.imgLupa
      .Left = mvarMe.Width - 680
      .Top = 100
      .Height = mvarMe.txtFiltrar.Height - 25
   End With
   With mvarMe.GrpEVENTO
      .Height = 1815
      .Top = mvarMe.ScaleHeight - .Height - 60
      .Left = 30
      .Width = mvarMe.ScaleWidth - 50
      
   End With
   mvarMe.GrpTarefa.Width = mvarMe.GrpEVENTO.Width - 240
   mvarMe.LblSTARTTIME.Left = mvarMe.GrpTarefa.Width - mvarMe.LblSTARTTIME.Width - 60
   mvarMe.LblEvento.Left = mvarMe.LblSTARTTIME.Left - mvarMe.LblEvento.Width - 60
   mvarMe.TxtOutros.Width = mvarMe.LblEvento.Left - mvarMe.TxtOutros.Left - 240
   
   mvarMe.TxtOutros.ShowBorder = True
   mvarMe.TxtOutros.Enabled = True
   mvarMe.TxtOutros.Locked = True
   mvarMe.TxtOutros.Top = mvarMe.LblEvento.Top
   mvarMe.TxtOutros.Height = mvarMe.GrpEVENTO.Height - mvarMe.TxtOutros.Top

   
   
   mvarMe.LblSituacao.Caption = "Cancelado"
   mvarMe.LblSituacao.Left = mvarMe.GrpTarefa.Width - mvarMe.LblSituacao.Width - 480
   
   With mvarMe.GrdTarefas
      .Left = 0
      .Top = mvarMe.SccConta.Top + mvarMe.SccConta.Height + 30
      .Width = mvarMe.ScaleWidth - 50
      .Height = mvarMe.GrpEVENTO.Top - .Top - 30
   End With


   'mvarMe.LblTitulo.Caption = mvarMe.TabConta.Item(mvarMe.TabConta.SelectedItem).Caption
   'mvarMe.LblTitulo.Left = mvarMe.TabConta.Width - mvarMe.LblTitulo.Width - 240

   Call DefineGrdAlinha

   'DoEvents
   mvarMe.GrpEVENTO.ZOrder 0
   'mvarMe.Refresh
End Sub
Private Sub mvarMe_Unload(Cancel As Integer)
  On Error Resume Next
End Sub
Private Function QueryGrid(Optional pWhere As String) As String
   Dim Sql As String

   Sql = "Select  '' [!], '' [Sit.] " & vbNewLine
   Sql = Sql & ", Cast(Convert(varchar(10), T.DTTAREFA , 103) As Datetime) [Data] " & vbNewLine
   Sql = Sql & ", Right('00'+CAST(T.IDTPTAREFA AS VARCHAR),2)+' - '+IsNull(P.DSCTPTAREFA,'Outras Tarefas') [Tipo]" & vbNewLine
   Sql = Sql & ", T.IDTAREFA [Cód.], T.TITULO [Tarefa]" & vbNewLine
   
   Sql = Sql & ", T.IDTAREFA, T.DTTAREFA, T.IDLOJA, T.IDEVENTO, T.IDCLIENTE" & vbNewLine
   Sql = Sql & ", T.SITTAREFA, T.PRIORIDADE, T.DSCTAREFA, T.FLGLEMBRETE, T.DTLEMBRETE, T.FLGDELETE " & vbNewLine
   Sql = Sql & ", C.NOME, C.TEL1, C.TEL2, C.FAX, C.EMAIL" & vbNewLine
   Sql = Sql & ", E.STARTDATETIME, E.FLGCONFIRMADO, E.FLGCANCELADO, E.FLGREMARCADO" & vbNewLine
   
   Sql = Sql & " From OTAREFAEVT T Left Join OTIPOTAREFA P On T.IDLOJA=P.IDLOJA And T.IDTPTAREFA=P.IDTPTAREFA" & vbNewLine
   Sql = Sql & " Left Join OCLIENTE C On T.IDLOJA=C.IDLOJA And T.IDCLIENTE=C.IDCLIENTE" & vbNewLine
   Sql = Sql & " Left Join OEVENTOAGENDA E On T.IDLOJA=E.IDLOJA And T.IDEVENTO=E.IDEVENTO" & vbNewLine
   Sql = Sql & " Where T.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And T.DTTAREFA>=CAST('01/09/2011 00:00:00' As DATETIME) " & vbNewLine ' & SqlDate(mvarSys.XDB.SysDate)
   If Not mvarTLPane Is Nothing Then
      If Not mvarTLPane.Pane Is Nothing Then
         Sql = Sql & " And ("
         If mvarTLPane.Pane.ChkN100.Value = 1 Or mvarTLPane.Pane.Chk100.Value = 1 Or mvarTLPane.Pane.ChkAndamento.Value = 1 Then
            Sql = Sql & " T.SITTAREFA in ("
            'Sql = Sql & " Or T.SITTAREFA in ("
            If mvarTLPane.Pane.ChkN100.Value = 1 Then Sql = Sql & "0, "
            If mvarTLPane.Pane.ChkAndamento.Value = 1 Then Sql = Sql & "5, "
            If mvarTLPane.Pane.Chk100.Value = 1 Then Sql = Sql & "10, "
            Sql = Mid(Sql, 1, Len(Sql) - 2)
            Sql = Sql & ")"
         Else
            Sql = Sql & " T.SITTAREFA in (-1)"
         End If
         If mvarTLPane.Pane.ChkDelete.Value = 1 Then
            Sql = Sql & " Or T.FLGDELETE=1"
         Else
            Sql = Sql & " And T.FLGDELETE=0"
         End If
         Sql = Sql & ")" & vbNewLine
         'Sql = Sql & " Or ("
         'Sql = Sql & " Cast(Convert(varchar(10), T.DTTAREFA , 103) As Datetime)=Cast(Convert(varchar(10), GetDate() , 103) As Datetime)"
         'Sql = Sql & " And T.FLGDELETE=0"
         'Sql = Sql & " )"
         
         Dim xDt As DatePicker
         Set xDt = mvarTLPane.Pane.DpiCalendario
         If xDt.Selection.BlocksCount > 0 Or mvarTLPane.Pane.TabPeriodo.SelectedItem = 1 Then
'            Sql = Sql & " And (T.DTTAREFA>=" & SqlDate(xDt.Selection.Blocks(0).DateBegin & " 00:00") & vbNewLine
'            Sql = Sql & " And  T.DTTAREFA<=" & SqlDate(xDt.Selection.Blocks(0).DateEnd & " 23:59") & ")" & vbNewLine
            Sql = Sql & " And (T.DTTAREFA>=" & SqlDate(Format(mvarTLPane.Pane.CmbDTIni.Value, "dd/mm/yyyy") & " 00:00") & vbNewLine
            Sql = Sql & " And  T.DTTAREFA<=" & SqlDate(Format(mvarTLPane.Pane.CmbDTFim.Value, "dd/mm/yyyy") & " 23:59") & ")" & vbNewLine
         Else
            Sql = Sql & " And (T.DTTAREFA>=" & SqlDate(Format(Now, "dd/mm/yyyy") & " 00:00") & vbNewLine
            Sql = Sql & " And  T.DTTAREFA<=" & SqlDate(Format(Now, "dd/mm/yyyy") & " 23:59") & ")" & vbNewLine
         End If
         
      End If
   End If
   Sql = Sql & " Order By T.DTTAREFA"

   QueryGrid = Sql
End Function
Public Sub PopulaGrdTarefas(Optional pWhere As String)
   Dim Sql As String
   Dim i As Long
   Dim nTopRow As Long
   Dim nRowSel As Long
   Dim nTabSel As Integer
   Dim oRecord As ReportRecord
   Dim oRptRow As ReportRow
   Dim CollExpanded As Collection
   
   Set CollExpanded = New Collection
   
   Screen.MousePointer = vbHourglass
   mvarMe.SccConta.Caption = "     Lista de Tarefas"
   If mvarTLPane.Pane.TabPeriodo.SelectedItem = 0 Then
      mvarMe.SccConta.Caption = mvarMe.SccConta.Caption & " (Dia: "
      mvarMe.SccConta.Caption = mvarMe.SccConta.Caption & Format(mvarTLPane.Pane.CmbDTIni.Value, "dd/mm/yy")
   Else
      mvarMe.SccConta.Caption = mvarMe.SccConta.Caption & " (Periodo: "
      mvarMe.SccConta.Caption = mvarMe.SccConta.Caption & Format(mvarTLPane.Pane.CmbDTIni.Value, "dd/mm/yy")
      mvarMe.SccConta.Caption = mvarMe.SccConta.Caption & " - " & Format(mvarTLPane.Pane.CmbDTFim.Value, "dd/mm/yy")
   End If
   mvarMe.SccConta.Caption = mvarMe.SccConta.Caption & ")"
   
   With mvarMe.GrdTarefas
      '.SetWatermarkBitmapFromFile mvarSys.ExePath & "DpilWater.bmp", 10
      '.WatermarkAlignment = xtpReportWatermarkPreserveRatio + xtpReportWatermarkCenter + xtpReportWatermarkVCenter
   
      nTopRow = .TopRowIndex
      If .SelectedRows.Count <= 0 Then
         nRowSel = -1
      Else
         nRowSel = .SelectedRows(0).Index
         Set oRecord = .SelectedRows(0).Record
      End If
      
      
      If .GroupsOrder.Count > 0 Then
         Dim GroupRow As ReportRow
         Dim bExpanded As Boolean
   
'         For Each GroupRow In .Rows
'            If GroupRow.GroupRow Then
'               If GroupRow.Expanded Then
'                  CollExpanded.Add GroupRow.Index, "k" & GroupRow.Index
'               End If
'            End If
'         Next
      End If
      '.Visible = False
      
      Sql = QueryGrid(pWhere)
      Call mvarSys.xDb.AbreTabela(Sql, RsLan)
      Call L_FillRCFromRS(RsLan, mvarMe.GrdTarefas)
      Call EscondeCampos
      If nRowSel < 0 Then
         nRowSel = RsLan.RecordCount
      End If

      
   '   Call DefineGrdIcones
      Call DefineGrdFiltro
      'If Not .Visible Then
      Call DefineGrdAlinha
      Call DefineGrupos
      If CollExpanded.Count > 0 Then
         For Each GroupRow In .Rows
            If GroupRow.GroupRow Then
               If ExisteItem(CollExpanded, "k" & GroupRow.Index) Then
                  GroupRow.Expanded = True
               Else
                  GroupRow.Expanded = False
               End If
            End If
         Next
      End If
      Set CollExpanded = Nothing
      'End If

      '.Populate
      'DoEvents

      If nRowSel >= .Rows.Count Then nRowSel = .Rows.Count - 1
      If nTopRow >= .Rows.Count Then nTopRow = .Rows.Count - 1

      .TopRowIndex = nTopRow
      If Not oRecord Is Nothing Then
         Set oRptRow = .Rows.FindRow(oRecord)
      End If
      If oRptRow Is Nothing Then
         If nRowSel >= 0 Then
            .Rows(nRowSel).Selected = True
            .Rows(nRowSel).EnsureVisible
         End If
      Else
         oRptRow.Selected = True
         oRptRow.EnsureVisible
      End If

      .Visible = True
      .Redraw
      If .Visible And .Enabled Then
         .SetFocus
         If oRptRow Is Nothing Then
            If nRowSel >= 0 Then
               .Rows(nRowSel).Selected = True
            End If
         Else
            oRptRow.Selected = True
         End If
      End If
   End With
   Screen.MousePointer = vbDefault
End Sub
Private Sub DefineGrupos()
   Dim GroupRow As ReportRow
   Dim bExpanded As Boolean
   
   With mvarMe.GrdTarefas
      
      If .GroupsOrder.Count <> 1 Then
         .GroupsOrder.DeleteAll
         '.GroupsOrder.Add .Columns(iCampo("Data"))
         .GroupsOrder.Add .Columns(iCampo("Tipo"))
      End If
      .Populate
'      For Each GroupRow In .Rows
'         If GroupRow.GroupRow Then
'            bExpanded = False
'            If GroupRow.Childs.Count > 0 Then
'               If GroupRow.Childs(0).GroupRow Then
'                  If GroupRow.Childs(0).Childs.Count > 0 Then
'                     If Not GroupRow.Childs(0).Childs(0).GroupRow Then
'                        bExpanded = (CDate(GroupRow.Childs(0).Childs(0).Record.Item(iCampo("Data")).Value) = CDate(Format(Now(), "dd/m/yyyy")))
'                     End If
'                  End If
'               Else
'                  If GroupRow.Childs.Count > 0 Then
'                     If Not GroupRow.Childs(0).GroupRow Then
'                        bExpanded = (CDate(GroupRow.Childs(0).Record.Item(iCampo("Data")).Value) = CDate(Format(Now(), "dd/m/yyyy")))
'                     End If
'                  End If
'               End If
'            End If
'            GroupRow.Expanded = bExpanded
'         End If
'      Next
   End With
End Sub
Private Function RegLan(sCampo As String, Optional pGrid As Object) As ReportRecordItem
   Dim nIndex As Integer
   Dim oRecI  As ReportRecordItem
   Dim oRec As ReportRecord

   If pGrid Is Nothing Then Set pGrid = mvarMe.GrdTarefas
   If pGrid.SelectedRows.Count <> 1 Then
      Set oRec = New ReportRecord
      Set oRecI = oRec.AddItem("")
      oRecI.Value = ""
      Set RegLan = oRecI
      Exit Function
   End If
   sCampo = UCase(sCampo)

   If xVal(GetTag(pGrid, "COUNT", 0)) = 0 Then
      Call SetTag(pGrid, "COUNT", pGrid.Columns.Count)
      For nIndex = 0 To pGrid.Columns.Count - 1
         Call SetTag(pGrid, UCase(mvarMe.GrdTarefas.Columns(nIndex).Caption), nIndex)
      Next
   End If
   nIndex = xVal(GetTag(pGrid, sCampo, 0))

   Set RegLan = pGrid.SelectedRows(0).Record(nIndex)
End Function
Private Sub LimpaGrd()
   Dim Sql As String

   On Error GoTo TrataErro

   Sql = QueryGrid()

   Set RsLan = Nothing
   Call mvarSys.xDb.AbreTabela(Sql, RsLan)
   Call FillRCFromRS(RsLan, mvarMe.GrdTarefas)
   mvarMe.GrdTarefas.Records.DeleteAll

   Call EscondeCampos
   Exit Sub
TrataErro:
   ShowError
End Sub
Private Sub EscondeCampos()
   Dim i As Integer

   Set iCampo = New Collection

   If RsLan Is Nothing Then Exit Sub
   If RsLan.State = 0 Then Exit Sub


   Call SetTag(mvarMe.GrdTarefas, "COUNT", RsLan.Fields.Count)
   For i = 0 To RsLan.Fields.Count - 1
      If ExisteItem(iCampo, RsLan.Fields(i).Name) Then
         mvarMe.GrdTarefas.Columns(i).Visible = False
      Else
         iCampo.Add i, RsLan.Fields(i).Name
         mvarMe.GrdTarefas.Columns(i).Visible = InArray(RsLan.Fields(i).Name, Array("!", "Sit.", "Tarefa"))
            
         Call SetTag(mvarMe.GrdTarefas, UCase(RsLan.Fields(i).Name), i)
      End If
   Next
End Sub
'Private Sub DefineGrdIcones()
'   Dim i As Integer
'   With mvarMe.GrdTarefas
'      .SetImageList mvarMe.ImageList1
'
'      For i = 0 To .Records.Count - 1
'         .Records(i).Item(1).Icon = 0
'      Next
'      '.Columns(1).Width = 20
'      '.Columns(1).Caption = ""
'      '.Columns(1).Icon = 1
'   End With
'End Sub
Private Sub DefineGrdFiltro()
   With mvarMe.GrdTarefas
      If mvarMe.txtFiltrar.Text = "" Or mvarMe.txtFiltrar.Text = StrPesquisa Then
         .FilterText = ""
      Else
         .FilterText = mvarMe.txtFiltrar.Text
      End If
   End With
End Sub
Private Sub DefineGrdAlinha()
   Dim vCol As ReportColumn
   With mvarMe.GrdTarefas
      For Each vCol In .Columns
         vCol.HeaderAlignment = xtpAlignmentLeft
         vCol.Alignment = vCol.HeaderAlignment
         vCol.Sortable = False
      Next
      
      .Columns(iCampo("Data")).HeaderAlignment = xtpAlignmentCenter
      .Columns(iCampo("Data")).Alignment = xtpAlignmentCenter
      .Columns(iCampo("!")).HeaderAlignment = xtpAlignmentCenter
      .Columns(iCampo("!")).Alignment = xtpAlignmentCenter
      '.Columns(iCampo("Tipo")).HeaderAlignment = xtpAlignmentRight
      '.Columns(iCampo("Tipo")).Alignment = xtpAlignmentRight
      
      .AutoColumnSizing = True
      .Columns(iCampo("!")).Width = 20
      .Columns(iCampo("Sit.")).Width = 20
      .Columns(iCampo("Data")).Width = 100
      .Columns(iCampo("Tipo")).Width = 30
      .Columns(iCampo("Cód.")).Width = 50
      .Columns(iCampo("Tarefa")).Width = 400

      '.Columns(iCampo("Tarefa")).Width = (.Width / Screen.TwipsPerPixelX) - 800
      '.Populate
   End With
End Sub
Public Sub L_FillRCFromRS(ByRef pRecordSet As Object, ByRef pReportControl As Object, Optional bDoEvents As Boolean = False, Optional bDeleteColumn As Boolean = False)  ' XtremeReportControl)
   Dim xField     As Object   ' ADODB.Field
   Dim xRecord    As XtremeReportControl.ReportRecord
   Dim xItem      As XtremeReportControl.ReportRecordItem
   Dim nIndex     As Integer
   Dim nMouse     As Integer
   Dim nType      As Integer

   nMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass

   pReportControl.Records.DeleteAll

   If Not bDeleteColumn Then
      'On Error Resume Next
      bDeleteColumn = (pReportControl.Columns.Count <> pRecordSet.Fields.Count)
   End If

   If bDeleteColumn Then
      pReportControl.Columns.DeleteAll
      pReportControl.Redraw
      For Each xField In pRecordSet.Fields
         pReportControl.Columns.Add nIndex, xField.Name, 100, True
         nIndex = nIndex + 1
      Next xField
   End If
  
   Do While Not pRecordSet.EOF
      Set xRecord = pReportControl.Records.Add()
      For nIndex = 0 To pRecordSet.Fields.Count - 1
         Set xItem = xRecord.AddItem("")
         With xItem
            If IsNull(pRecordSet.Collect(nIndex)) Then
               .Value = Empty
            Else
               nType = pRecordSet.Fields(nIndex).Type
               If nType = adCurrency Or nType = adDecimal Or nType = adDouble Or nType = adNumeric Then
                  .Value = pRecordSet.Collect(nIndex)
                  If .Value < 0 Then
                     .Caption = "(" & Mid(ValBr(pRecordSet.Collect(nIndex)), 2) & ")"
                     .ForeColor = vbRed
                  Else
                     .Caption = ValBr(pRecordSet.Collect(nIndex))
                  End If
               Else
                  .Value = pRecordSet.Collect(nIndex)
               End If
               
               '*Private
               If pRecordSet("FLGDELETE") = 1 Then
                  .ForeColor = vbRed
               ElseIf pRecordSet("SITTAREFA") = 10 Then
                  .ForeColor = RGB(128, 128, 128)
               Else
                  .ForeColor = vbBlack
               End If
               If pRecordSet.Fields(nIndex).Name = "Sit." Then
                  If pRecordSet("FLGDELETE") = 1 Then
                     .Icon = 2
                  Else
                     Select Case pRecordSet("SITTAREFA").Value
                        Case 0   'Item.Icon = 1
                        Case 5:  .Icon = 3
                        Case 10: .Icon = 1
                     End Select
                  
                  End If
               End If
            End If
         End With
      Next
      If bDoEvents Then DoEvents
      pRecordSet.MoveNext
   Loop
   
   '* Private
   pReportControl.Columns(0).Icon = 5
   pReportControl.Columns(0).Caption = ""

   
   pReportControl.Populate
   Screen.MousePointer = nMouse
End Sub
Public Sub EditarObjeto(Optional pId As Integer)
   Dim TlCadastro As New TL_CadTarefa
   If pId = 0 Then
      If mvarMe.GrdTarefas.SelectedRows.Count = 0 Then
         Call ExibirAviso("Selecione uma Tarefa!", "Tarefas")
         Exit Sub
      End If
      If mvarMe.GrdTarefas.SelectedRows(0).Record Is Nothing Then
         'Call ExibirAviso("Selecione tarefa.", "Tarefas")
         Exit Sub
      End If
      pId = mvarMe.GrdTarefas.SelectedRows(0).Record.Item(iCampo("IDTAREFA")).Value
   End If
   With TlCadastro
      Set .Sys = mvarSys
      .IDLOJA = mvarIDLOJA
      .IDTAREFA = pId
      Set .TLCal = mvarTLCal
      
      Call .Show
      If .IsDirt Then
         Call PopulaGrdTarefas
      End If
   End With
   Set TlCadastro = Nothing
   
   mvarMe.GrdTarefas.Redraw
   mvarMe.GrdTarefas.SetFocus
   Call mvarMe_GrdTarefasSelectionChanged
End Sub

