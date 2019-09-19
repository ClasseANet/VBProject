VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_PaneTarefa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarPane  As FrmPaneTarefa
Attribute mvarPane.VB_VarHelpID = -1
Dim TLTarefas As TL_Tarefas

Private Enum TaskItem
   NovaTarefa = 1
   EditTarefa = 2
   ExclTarefa = 3
   ImprTarefa = 4
End Enum
Dim RsDatas As Recordset

Private mvarSys      As Object
Private mvarItemCad  As Integer
Private mvarTLCal    As Object
Public Property Set TLCal(ByVal vData As Object)
   Set mvarTLCal = vData
End Property
Public Property Get TLCal() As Object
   Set TLCal = mvarTLCal
End Property
Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
    Set TLTarefas.Sys = mvarSys
    
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Property Get Pane() As Object
    Set Pane = mvarPane
End Property
Public Sub Show(Optional Reload As Boolean = False)
   If Not TLTarefas Is Nothing Then
      Set TLTarefas.Sys = mvarSys
      Set TLTarefas.TLCal = mvarTLCal
      Set TLTarefas.TLPane = Me
      TLTarefas.Show Reload
   End If
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarPane
   Set mvarPane = Nothing
End Sub

Private Sub mvarPane_Activate()
  ' Call mvarPane_Load
End Sub
Private Sub mvarPane_CheckClick()
   TLTarefas.PopulaGrdTarefas
End Sub
Private Sub mvarPane_CmbDTFimChange()
'   Call TLTarefas.PopulaGrdTarefas
End Sub
Private Sub mvarPane_CmbDTIniChange()
'   Call TLTarefas.PopulaGrdTarefas
End Sub
Private Sub mvarPane_CmdPeriodoClick()
   Call TLTarefas.PopulaGrdTarefas
End Sub
Private Sub mvarPane_CommandBarsExecute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   'If Control.Category = "Popup_IDCONTA" Then
   '   If TypeName(TLTarefas) = "TL_MovCC" Then
   '      mvarSys.Propriedades("IDCONTA") = Control.Id
   '      TLTarefas.IDTAREFA = Control.Id
   '      TLTarefas.PopulaTela
   '      mvarPane.SccConta2.Caption = "Minhas Tarefas"
   '   End If
   'End If
End Sub
Private Sub mvarPane_DpiCalendarioDayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
   If Weekday(Day) = vbSunday Then
      Metrics.ForeColor = vbRed
   End If
   
   If DatePart("d", Day) = 1 Then
      Call PopulaRsDatas
   End If
   If RsDatas Is Nothing Then
      Metrics.Font.Bold = False
   Else
      Call RsDatas.Find("DTTAREFA=" & Day, Start:=1)
      If (RsDatas.BOF = True) Or (RsDatas.EOF = True) Then
         Metrics.Font.Bold = False
      Else
         Metrics.Font.Bold = True
      End If
   End If
   
   
'   Sql = "Select Count(*) "
'   Sql = Sql & " From OTAREFAEVT"
'   Sql = Sql & " Where Cast(Convert(char, DTTAREFA, 103) as DateTime) =" & SqlDate(Day)
'   Sql = Sql & " And SITTAREFA In (0,5)"
'   If mvarSys.xDb.Abretabela(Sql) Then
'      If mvarSys.xDb.RsAux(0) > 0 Then
'         Metrics.Font.Bold = True
'      Else
'         Metrics.Font.Bold = False
'      End If
'   Else
'      Metrics.Font.Bold = False
'   End If
End Sub

Private Sub mvarPane_DpiCalendarioSelectionChanged()
   mvarPane.CmbDTIni.Value = mvarPane.DpiCalendario.Selection.Blocks(0).DateBegin
   mvarPane.CmbDTFim.Value = mvarPane.DpiCalendario.Selection.Blocks(0).DateEnd
   TLTarefas.PopulaGrdTarefas
'   With mvarPane.DpiCalendario
'   End With
   'xDt.Selection.Blocks(0).DateBegin
End Sub

Private Sub mvarPane_Load()
   Call PopulaRsDatas
   With mvarPane
      .CommandBars.DeleteAll
      .ChkN100.Value = xtpChecked
      .ChkAndamento.Value = xtpChecked
      .Chk100.Value = xtpUnchecked
      .ChkDelete.Value = xtpUnchecked
   
      .TabPeriodo.Move 0, 0, 2660, 2800
      With .DpiCalendario
         .BorderStyle = xtpDatePickerBorderNone
         .ShowNoneButton = False
         .FirstDayOfWeek = 1
         .AskDayMetrics = True
         .BoldDaysPerIdleStep = 3
         .BoldDaysIdleStepTime_ms = 60
         .AutoSizeRowCol = True
        
         .Move 60, 60, .Width, 2200
         .ShowNonMonthDays = False
       End With
      '.FraCalendario.Move 360, 2200, .Width
      .FraCalendario.Height = 2800
      
   End With
   Call CreateTaskPanel
End Sub
Private Sub mvarPane_Resize()
   With mvarPane
      .SccConta.Move 0, 0, .ScaleWidth, 420
      .SccConta2.Move 0, .SccConta.Height, .SccConta.Width, 285
   
      .Picture1.Move 0, .SccConta2.Top + .SccConta2.Height, .ScaleWidth, .ScaleHeight - .SccConta.Height
      .wndTaskPanel.Move 0, 0, .Picture1.ScaleWidth, .Picture1.ScaleHeight
      
      If .Width - 600 > 2500 Then
         .DpiCalendario.Width = 2500
      Else
         .DpiCalendario.Width = .Width - 720
      End If
      
      mvarPane.DpiCalendario.ShowNoneButton = False
      
   End With
End Sub

Private Sub mvarPane_TabPeriodoSelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   Dim dIni As Date
   Dim dFim As Date
   If Item.Index = 0 Then
      dIni = mvarPane.CmbDTIni.Value
      dFim = mvarPane.CmbDTFim.Value
      If mvarPane.DpiCalendario.Selection.BlocksCount > 0 Then
         mvarPane.CmbDTIni.Value = mvarPane.DpiCalendario.Selection.Blocks(0).DateBegin
         mvarPane.CmbDTFim.Value = mvarPane.DpiCalendario.Selection.Blocks(0).DateEnd
      Else
         mvarPane.CmbDTIni.Value = Format(Now, "dd/mm/yyyy")
         mvarPane.CmbDTFim.Value = Format(Now, "dd/mm/yyyy")
      End If
      If dIni <> mvarPane.CmbDTIni.Value Or dFim <> mvarPane.CmbDTFim.Value Then
         Call TLTarefas.PopulaGrdTarefas
      End If
   End If
End Sub

Private Sub mvarPane_wndTaskPanelItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
   Dim nRow As Long
   Screen.MousePointer = vbHourglass
   
   Dim oCAD As TL_CadTarefa
   
   
   Select Case Item.Id
      Case 1:       'ADD CONTATO
         Set oCAD = New TL_CadTarefa
         With oCAD
            Set .Sys = mvarSys
            Call .Show
            If .IsDirt = True Then
               If TLTarefas.frm.GrdTarefas.SelectedRows.Count = 1 Then
                  'nRow = TLTarefas.Frm.GrdTarefas.SelectedRows(0).Index
                  nRow = TLTarefas.frm.GrdTarefas.FocusedRow.Index
                  TLTarefas.frm.GrdTarefas.Rows(nRow).EnsureVisible
                  TLTarefas.frm.GrdTarefas.Rows(nRow).Selected = True
                End If
               TLTarefas.PopulaGrdTarefas
            End If
         End With
         Set oCAD = Nothing
      Case 2:
         Call TLTarefas.EditarObjeto
         
      Case 3:       'DEL CONTATO
         If TLTarefas.frm.GrdTarefa.SelectedRows.Count > 0 Then
            Set oCAD = New TL_CadTarefa
            With oCAD
               Set .Sys = mvarSys
               .IDTAREFA = TLTarefas.frm.GrdTarefa.SelectedRows(0).Record.Item(0).Value
               Call .ExcluirTarefa
               If .IsDirt = True Then
                  TLTarefas.PopulaGrdTarefas
               End If
            End With
            Set oCAD = Nothing
         End If
      Case 4:       '
      Case 5:       '
      Case 6:       '
      
   End Select
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub Class_Initialize()
   Set mvarPane = New FrmPaneTarefa
   Set TLTarefas = New TL_Tarefas
End Sub
Private Sub CreateTaskPanel()
   Dim Group As TaskPanelGroup
   Dim Item As TaskPanelGroupItem
   
   mvarPane.wndTaskPanel.SetImageList mvarPane.imlTaskPanelIcons
   
   Set Group = mvarPane.wndTaskPanel.Groups.Add(0, "Operações")
   Group.Items.Add TaskItem.NovaTarefa, "Adicionar nova tarefa", xtpTaskItemTypeLink, 4
   Group.Items.Add TaskItem.EditTarefa, "Editar tarefa selecionada", xtpTaskItemTypeLink, 5
   Set Item = Group.Items.Add(TaskItem.ExclTarefa, "Excluir tarefa selecionada", xtpTaskItemTypeLink, 6)
   Item.Enabled = False
   Group.Expandable = False
    
   Set Group = mvarPane.wndTaskPanel.Groups.Add(0, "Calendário")
   Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
   Set Item.Control = mvarPane.FraCalendario
   mvarPane.FraCalendario.BackColor = Item.BackColor
   Group.Expanded = True
   Group.Expandable = False
   'Group.MinimumClientHeight = 1
    
   Set Group = mvarPane.wndTaskPanel.Groups.Add(0, "Filtros")
   Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
    Set Item.Control = mvarPane.FraFiltro
    mvarPane.FraFiltro.BackColor = Item.BackColor
    mvarPane.ChkN100.BackColor = Item.BackColor
    mvarPane.Chk100.BackColor = Item.BackColor
    mvarPane.ChkAndamento.BackColor = Item.BackColor
    mvarPane.ChkDelete.BackColor = Item.BackColor
'   Group.Items.Add ID_TASKITEM_NEWFOLDER, "Visualizar tratamentos", xtpTaskItemTypeControl, 1
'   Group.Items.Add ID_TASKITEM_PUBLISH, "Últimas comunicações", xtpTaskItemTypeLink, 2
'   Group.Items.Add ID_TASKITEM_SHARE, "Histório de compra do creme", xtpTaskItemTypeLink, 3
   Group.Expanded = True
    
'   Set Group = mvarPane.wndTaskPanel.Groups.Add(0, "Busca específica")
'   Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
'   Set Item.Control = mvarPane.fraBuscaDetalhada
'   mvarPane.fraBuscaDetalhada.BackColor = Item.BackColor
'   Group.Expanded = False
    
    mvarPane.wndTaskPanel.Reposition
End Sub

Private Sub PopulaRsDatas()
   Dim Sql As String
   Sql = "Select Distinct Cast(Convert(varchar, DTTAREFA, 103) as DateTime) [DTTAREFA]"
   Sql = Sql & " From OTAREFAEVT"
   Sql = Sql & " Where FLGDELETE = 0"
   Sql = Sql & " And SITTAREFA In (0,5)"
   Sql = Sql & " And IDLOJA=" & mvarSys.Propriedades("IDLOJA")
   Sql = Sql & " Group By Cast(Convert(varchar, DTTAREFA, 103) as DateTime)"
   Sql = Sql & " Having Count(*) >0"
   If Not mvarSys.xDb.AbreTabela(Sql, RsDatas) Then
      Set RsDatas = Nothing
   End If
End Sub
