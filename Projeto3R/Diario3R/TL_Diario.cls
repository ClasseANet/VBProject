VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Diario"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const StrPesquisa = "Pesquisar Ocorrências"
Private WithEvents mvarMe     As FrmDiario
Attribute mvarMe.VB_VarHelpID = -1
Private WithEvents mvarCmmdBars  As XtremeCommandBars.CommandBars
Attribute mvarCmmdBars.VB_VarHelpID = -1
Private mvarTLPane As TL_PaneDiario
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

Dim sDia       As String
Dim sTexto     As String
'Private bNovo As Boolean
Private RsLan        As Object
Private mvarSys      As Object
Private mvarIDLOJA   As Integer
Private mvarDTDIARIO As String
Private mvarTLCal    As Object
Public Property Set TLPane(ByVal vData As TL_PaneDiario)
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
Public Property Let DTDIARIO(ByVal vData As String)
   mvarDTDIARIO = vData
End Property
Public Property Get DTDIARIO() As String
   DTDIARIO = mvarDTDIARIO
End Property
Public Sub Show(Optional Reload As Boolean = False)
   '* Carregar e Acoplar Form no MDI
   If Reload Then
      Set Sys = Sys
      Unload mvarMe
   End If
   '* Acoplar Form no MDI
   Load mvarMe
   Call AcoplarForm(mvarMe, 2, mvarSys)
   mvarMe.Show

   '* Exibir Menu no MDI
   If Not mvarSys.MDI.CommandBars.FindControl(, Mnu_Tarefas, False) Is Nothing Then
      mvarSys.MDI.CommandBars.FindControl(, Mnu_Tarefas).Visible = True
   End If

   '* Definir foco
   With mvarMe.GrdDiario
      DoEvents
      If .Visible And .Enabled Then
         'Call MontaTela
         Call PopulaGrdDiario
         .SetFocus
      End If
   End With
End Sub
Public Sub Hide()
   mvarMe.Hide
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmDiario
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub

Private Sub mvarCmmdBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   Dim n As Object
   
   DoEvents
   Select Case UCase(Control.Category)
      Case "TAREFA"
         Select Case Control.Id
            Case Mnu_TarSitNormal, Mnu_TarSitAndamento, Mnu_TarSitConcluido
               Call AlterarSitTarefa(Control.Id - Mnu_TarSitNormal)
               
            Case Mnu_TarEditarNova
            
         End Select
   End Select
End Sub
Private Sub AlterarSitTarefa(pIndex As Long)
   Dim TbTarefa As Object
   Dim nIdTarefa As Long
   Dim i As Integer
   Dim xItem As ReportRecordItem

   nIdTarefa = xVal(GetTag(mvarMe.GrdDiario, "IDTAREFA", 0))
   If nIdTarefa <> 0 Then
      Set TbTarefa = CriarObjeto("BANCO_3R.TB_OTAREFAEVT")
      Set TbTarefa.xdb = mvarSys.xdb
      If TbTarefa.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDTAREFA:=nIdTarefa) Then
         Select Case pIndex
            Case 0: TbTarefa.SITTAREFA = 0 '"Não Iniciada", 0
            Case 1: TbTarefa.SITTAREFA = 5 '"Em Andamento", 5
            Case 2: TbTarefa.SITTAREFA = 10 '"Concluída", 10
         End Select
         If TbTarefa.IsDirt Then
            If TbTarefa.Salvar Then
               Call mvarMe_KeyUp(vbKeyF5, 0)
            
               On Error Resume Next
'               mvarMe.GrdDiario.FocusedRow.Selected = True
'               'mvarMe.GrdDiario.SelectedRows(0).Record
'               i = mvarMe.GrdDiario.SelectedRows.Count - 1
'               Set xItem = mvarMe.GrdDiario.SelectedRows(i).Record.Item(0)
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
'                  mvarMe.GrdDiario.Redraw
'               End If
            End If
         End If
      End If
   End If
End Sub
Private Sub mvarMe_Activate()
'   Call MontaTela
'   Call PopulaTela
   If GetTag(mvarMe, "LOAD", 1) = 1 Then
      Call SetTag(mvarMe, "LOAD", 0)
      mvarMe.GrdDiario.Visible = False
      Call mvarMe_GrdDiariosSelectionChanged
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarMe_GrdDiariosBeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
   
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
Private Sub mvarMe_GrdDiariosMouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
   Dim Popup     As CommandBar
   Dim xItem As ReportRecordItem
   Dim xRow As ReportRow
   
   Call SetTag(mvarMe.GrdDiario, "IDTAREFA", 0)
   If Button = 2 Then
      Set xItem = mvarMe.GrdDiario.HitTest(x, y).Item
      Set xRow = mvarMe.GrdDiario.HitTest(x, y).Row
      If Not xItem Is Nothing Then
         mvarMe.GrdDiario.FocusedRow.Selected = False
         xRow.Selected = True
         Call SetTag(mvarMe.GrdDiario, "IDTAREFA", xItem.Record(iCampo("IDTAREFA")).Value)
         Call SetTag(mvarMe.GrdDiario, "ROW", xRow.Index)
         Set Popup = mvarSys.MDI.CommandBars.ContextMenus.Find(Pop_TarSit)
         If Not Popup Is Nothing Then
            Popup.ShowPopup
         End If
      End If
   End If
End Sub
Private Sub mvarMe_GrdDiariosSelectionChanged()
   DoEvents
   If Not mvarMe.GrdDiario.Visible Then Exit Sub
   Call PopulaTexto
   mvarMe.GrdDiario.SetFocus
End Sub
Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim xIDTarefa As Long
   Dim iRow As Long
   Dim xRow As ReportRow
   
   If Not mvarMe.GrdDiario.Visible Then Exit Sub
   DoEvents
   If KeyCode = vbKeyF5 Then
      If mvarMe.GrdDiario.SelectedRows.Count > 0 Then
         If Not mvarMe.GrdDiario.SelectedRows(0).Record Is Nothing Then
            xIDTarefa = mvarMe.GrdDiario.SelectedRows(0).Record.Item(iCampo("IDTAREFA")).Value
         End If
      End If
      Call PopulaGrdDiario
      
      If xIDTarefa > 0 Then
         For iRow = 0 To mvarMe.GrdDiario.Records.Count - 1
            If mvarMe.GrdDiario.Records(iRow).Item(iCampo("IDTAREFA")).Value = xIDTarefa Then
               mvarMe.GrdDiario.Records(iRow).Expanded = True
               Set xRow = mvarMe.GrdDiario.Rows.FindRowInTree(mvarMe.GrdDiario.Records(iRow))
               Set xRow = xRow.ParentRow
               While Not xRow Is Nothing
                  xRow.Expanded = True
                  Set xRow = xRow.ParentRow
               Wend
               Set xRow = mvarMe.GrdDiario.Rows.FindRowInTree(mvarMe.GrdDiario.Records(iRow))
               xRow.EnsureVisible
               xRow.Selected = True
               mvarMe.GrdDiario.Redraw
               Exit For
            End If
         Next
      End If
   ElseIf KeyCode = vbKeyDelete Then
      If mvarMe.ActiveControl Is mvarMe.GrdDiario Then
         'eXCLUI
      End If

   ElseIf KeyCode = vbKeyEscape Then
      If mvarMe.txtFiltrar.Text <> StrPesquisa Then
         mvarMe.txtFiltrar.Text = StrPesquisa
         Call PopulaGrdDiario
      End If

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
   Call PopulaGrdDiario
End Sub
Private Sub MontarTela()
   'ReportControlGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   With mvarMe
      .txtFiltrar.Visible = False
      .imgLupa.Visible = False
      .txtFiltrar.Text = StrPesquisa
      .SccConta.VisualTheme = xtpShortcutThemeOffice2003
      .SccConta.GradientHorizontal = False
      With .GrdDiario
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
      AddButtonBar .Controls, Mnu_TarMail, "eMail", True, xtpControlButton, Category:="TAREFA"
      
   End With
   ContextMenu.Controls.CreateOriginalControls
   
End Sub
Private Sub mvarMe_Rezise()
   On Error Resume Next

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
   With mvarMe.TxtDSCDIARIO
      .Left = 0
      .Top = mvarMe.SccConta.Top + mvarMe.SccConta.Height + 30
      .Width = mvarMe.ScaleWidth - 50
      .Height = mvarMe.ScaleHeight - 50
   End With
   With mvarMe.GrdDiario
      .Left = 0
      .Top = mvarMe.SccConta.Top + mvarMe.SccConta.Height + 30
      .Width = mvarMe.ScaleWidth - 50
      .Height = (mvarMe.ScaleHeight - 50) / 2
   End With
   
   Call DefineGrdAlinha
   mvarMe.TxtDSCDIARIO.ZOrder 0

End Sub

Private Sub mvarMe_Timer1()
   If sTexto <> mvarMe.TxtDSCDIARIO.Text Then
      Call SalvarOcorrencia
   End If
End Sub
Private Sub mvarMe_TxtDSCDIARIOChange()
   mvarMe.Timer1.Enabled = True
End Sub
Private Sub mvarMe_TxtDSCDIARIOLostFocus()
   Call SalvarOcorrencia
End Sub
Public Sub SalvarOcorrencia()
   Dim Sql As String
   Screen.MousePointer = vbHourglass
   If sTexto <> mvarMe.TxtDSCDIARIO.Text Then
      If IsDate(sDia) Then
         If CDate(sDia) = mvarSys.xdb.SysDate() Then
            Sql = Sql & "Set DateFormat 'DMY';" & vbNewLine
            Sql = Sql & "If Exists(Select IDLOJA, DTDIARIO, DSCDIARIO, ALTERSTAMP" & vbNewLine
            Sql = Sql & ", TIMESTAMP" & vbNewLine
            Sql = Sql & " From ODIARIO" & vbNewLine
            Sql = Sql & " Where" & vbNewLine
            Sql = Sql & " IDLOJA = " & mvarIDLOJA & vbNewLine
            Sql = Sql & " And DTDIARIO =" & SqlDate(sDia) & ")" & vbNewLine
            Sql = Sql & " Update ODIARIO Set" & vbNewLine
            Sql = Sql & " IDLOJA = " & mvarIDLOJA & vbNewLine
            Sql = Sql & " , DTDIARIO   = " & SqlDate(sDia) & vbNewLine
            Sql = Sql & " , DSCDIARIO  = " & SqlStr(mvarMe.TxtDSCDIARIO.Text) & vbNewLine
            Sql = Sql & ", ALTERSTAMP = 1" & vbNewLine
            Sql = Sql & ", TIMESTAMP  = GetDate()" & vbNewLine
            Sql = Sql & " Output Inserted.*" & vbNewLine
            Sql = Sql & " Where" & vbNewLine
            Sql = Sql & " IDLOJA = " & mvarIDLOJA & vbNewLine
            Sql = Sql & " And DTDIARIO = " & SqlDate(sDia) & vbNewLine
            Sql = Sql & " Else" & vbNewLine
            Sql = Sql & " Insert Into ODIARIO (IDLOJA, DTDIARIO, DSCDIARIO, ALTERSTAMP" & vbNewLine
            Sql = Sql & ", TIMESTAMP)" & vbNewLine
            Sql = Sql & " Values" & vbNewLine
            Sql = Sql & " (" & mvarIDLOJA & vbNewLine
            Sql = Sql & ", " & SqlDate(sDia) & vbNewLine
            Sql = Sql & ", " & SqlStr(mvarMe.TxtDSCDIARIO.Text) & vbNewLine
            Sql = Sql & ", 1" & vbNewLine
            Sql = Sql & ", GetDate()" & vbNewLine
            Sql = Sql & ")" & vbNewLine
            If mvarSys.xdb.Executa(Sql) Then
               sTexto = mvarMe.TxtDSCDIARIO.Text
               mvarMe.Timer1.Enabled = False
            End If
         End If
      End If
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_Unload(Cancel As Integer)
  On Error Resume Next
End Sub
Private Function QueryGrid(Optional pWhere As String) As String
   Dim Sql As String

   Sql = "Select D.IDLOJA, Cast(Convert(varchar, DTDIARIO, 103) as DateTime) [Data], D.DSCDIARIO [Ocorrência]" & vbNewLine
   Sql = Sql & " From ODIARIO D" & vbNewLine
   Sql = Sql & " Where IDLOJA=" & mvarSys.IDLOJA & vbNewLine
   If Not mvarTLPane Is Nothing Then
      Dim xDt As DatePicker
      Set xDt = mvarTLPane.Pane.DpiCalendario
      If xDt.Selection.BlocksCount > 0 Then
         sDia = xDt.Selection.Blocks(0).DateBegin
      Else
         sDia = Format(Now, "dd/mm/yyyy")
      End If
      Sql = Sql & " And (D.DTDIARIO>=" & SqlDate(sDia & " 00:00") & vbNewLine
      Sql = Sql & " And  D.DTDIARIO<=" & SqlDate(sDia & " 23:59") & ")" & vbNewLine
   End If
   Sql = Sql & " Order By D.DTDIARIO"

   QueryGrid = Sql
End Function
Public Sub PopulaGrdDiario(Optional pWhere As String)
   Dim Sql As String
   Dim nTopRow As Long
   Dim nRowSel As Long
   Dim oRecord As ReportRecord
   Dim oRptRow As ReportRow
  
   Screen.MousePointer = vbHourglass
   
   sTexto = ""
   Sql = QueryGrid(pWhere)
   If mvarSys.xdb.Abretabela(Sql, RsLan) Then
      sTexto = RsLan("Ocorrência")
   End If
   mvarMe.TxtDSCDIARIO.Text = sTexto
   If IsDate(sDia) Then
      mvarMe.TxtDSCDIARIO.Locked = (CDate(sDia) <> mvarSys.xdb.SysDate())
      mvarMe.SccConta.Caption = "Ocorrência: " & sDia & " - " & DiaDaSemana(CDate(sDia))
   End If
   
   With mvarMe.GrdDiario
'      Sql = QueryGrid(pWhere)
'      If mvarSys.xDb.Abretabela(Sql, RsLan) Then
'         mvarMe.TxtDSCDIARIO.Text = RsLan("")
'      Else
'         mvarMe.TxtDSCDIARIO.Text = RsLan("")
'      End If
'      Call FillRCFromRS(RsLan, mvarMe.GrdDiario)
'      Call EscondeCampos
'      If nRowSel < 0 Then
'         nRowSel = RsLan.RecordCount
'      End If
'      Call DefineGrdFiltro
'
'      'If Not .Visible Then
'      Call DefineGrdAlinha
'      'End If
'
'      '.Populate
'      'DoEvents
'
'      If nRowSel >= .Rows.Count Then nRowSel = .Rows.Count - 1
'      If nTopRow >= .Rows.Count Then nTopRow = .Rows.Count - 1
'
'      .TopRowIndex = nTopRow
'      If Not oRecord Is Nothing Then
'         Set oRptRow = .Rows.FindRow(oRecord)
'      End If
'      If oRptRow Is Nothing Then
'         If nRowSel >= 0 Then
'            .Rows(nRowSel).Selected = True
'            .Rows(nRowSel).EnsureVisible
'         End If
'      Else
'         oRptRow.Selected = True
'         oRptRow.EnsureVisible
'      End If
'
'      .Visible = True
'      .Redraw
'      If .Visible And .Enabled Then
'         .SetFocus
'         If oRptRow Is Nothing Then
'            If nRowSel >= 0 Then
'               .Rows(nRowSel).Selected = True
'            End If
'         Else
'            oRptRow.Selected = True
'         End If
'      End If
'      Call PopulaTexto
   End With

   Screen.MousePointer = vbDefault
End Sub
Private Sub PopulaTexto()
   Dim xItem As ReportRecordItem
   
   If Not mvarMe.GrdDiario.Visible Then Exit Sub
   If mvarMe.GrdDiario.Rows.Count > 0 Then
      Set xItem = mvarMe.GrdDiario.SelectedRows(1).Record.Item(0)
      If Not xItem Is Nothing Then
         mvarMe.TxtDSCDIARIO.Text = xItem.Record(iCampo("DSCDIARIO")).Value
      End If
   End If
End Sub
Private Sub LimpaGrd()
   Dim Sql As String

   On Error GoTo TrataErro

   Sql = QueryGrid()

   Set RsLan = Nothing
   Call mvarSys.xdb.Abretabela(Sql, RsLan)
   Call FillRCFromRS(RsLan, mvarMe.GrdDiario)
   mvarMe.GrdDiario.Records.DeleteAll
   mvarMe.TxtDSCDIARIO.Text = ""

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


   Call SetTag(mvarMe.GrdDiario, "COUNT", RsLan.Fields.Count)
   For i = 0 To RsLan.Fields.Count - 1
      If ExisteItem(iCampo, RsLan.Fields(i).Name) Then
         mvarMe.GrdDiario.Columns(i).Visible = False
      Else
         iCampo.Add i, RsLan.Fields(i).Name
         mvarMe.GrdDiario.Columns(i).Visible = InArray(RsLan.Fields(i).Name, Array("!", "Sit.", "Tarefa"))
            
         Call SetTag(mvarMe.GrdDiario, UCase(RsLan.Fields(i).Name), i)
      End If
   Next
End Sub
Private Sub DefineGrdFiltro()
   With mvarMe.GrdDiario
      If mvarMe.txtFiltrar.Text = "" Or mvarMe.txtFiltrar.Text = StrPesquisa Then
         .FilterText = ""
      Else
         .FilterText = mvarMe.txtFiltrar.Text
      End If
   End With
End Sub
Private Sub DefineGrdAlinha()
   Dim vCol As ReportColumn
   With mvarMe.GrdDiario
      For Each vCol In .Columns
         vCol.HeaderAlignment = xtpAlignmentLeft
         vCol.Alignment = vCol.HeaderAlignment
         vCol.Sortable = False
      Next
      
      .Columns(iCampo("Data")).HeaderAlignment = xtpAlignmentLeft
      .Columns(iCampo("Data")).Alignment = xtpAlignmentLeft
      .Columns(iCampo("Ocorrência")).HeaderAlignment = xtpAlignmentLeft
      .Columns(iCampo("Ocorrência")).Alignment = xtpAlignmentLeft
      .AutoColumnSizing = True
      .Columns(iCampo("Data")).Width = 100
   End With
End Sub


