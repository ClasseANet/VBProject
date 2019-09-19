VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_PaneDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarPane  As FrmPaneDiario
Attribute mvarPane.VB_VarHelpID = -1
Dim TLPrincipal As TL_Diario

Dim RsDatas As Recordset

Private mvarSys      As Object
Private mvarTLCal    As Object
Public Property Set TLCal(ByVal vData As Object)
   Set mvarTLCal = vData
End Property
Public Property Get TLCal() As Object
   Set TLCal = mvarTLCal
End Property
Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
    Set TLPrincipal.Sys = mvarSys
    
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Property Get Pane() As Object
    Set Pane = mvarPane
End Property
Public Sub Show(Optional Reload As Boolean = False)
   If Not TLPrincipal Is Nothing Then
      With TLPrincipal
         Set .Sys = mvarSys
         Set .TLCal = mvarTLCal
         Set .TLPane = Me
         .Show Reload
      End With
   End If
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarPane
   Set mvarPane = Nothing
End Sub
Private Sub mvarPane_CheckClick()
   TLPrincipal.PopulaGrdDiario
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
      Call RsDatas.Find("DTDIARIO=" & Day, Start:=1)
      If (RsDatas.BOF = True) Or (RsDatas.EOF = True) Then
         Metrics.Font.Bold = False
      Else
         Metrics.Font.Bold = True
      End If
   End If
End Sub

Private Sub mvarPane_DpiCalendarioSelectionChanged()
   TLPrincipal.PopulaGrdDiario
'   With mvarPane.DpiCalendario
'   End With
   'xDt.Selection.Blocks(0).DateBegin
End Sub

Private Sub mvarPane_Load()
   Call PopulaRsDatas
   With mvarPane
      .CommandBars.DeleteAll
   
      With .DpiCalendario
         .BorderStyle = xtpDatePickerBorderNone
         .ShowNoneButton = False
         .FirstDayOfWeek = 1
         .AskDayMetrics = True
         .BoldDaysPerIdleStep = 3
         .BoldDaysIdleStepTime_ms = 60
         .AutoSizeRowCol = True
        
         .Move 0, 0, .Width, 2400
         .ShowNonMonthDays = False
       End With
      .FraCalendario.Height = 2400
      
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
      
      .DpiCalendario.ShowNoneButton = False
      
   End With
End Sub

Private Sub mvarPane_wndTaskPanelItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
   Dim nRow As Long
   Screen.MousePointer = vbHourglass
   
   Dim oCAD As TL_Diario
   
   
   Select Case Item.Id
      Case 1:       'ADD CONTATO
      Case 2:
         'Call TLPrincipal.EditarObjeto
         
      Case 3:       'DEL CONTATO
         If TLPrincipal.frm.GrdDiario.SelectedRows.Count > 0 Then
         End If
      Case 4:       '
      Case 5:       '
      Case 6:       '
      
   End Select
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub Class_Initialize()
   Set mvarPane = New FrmPaneDiario
   Set TLPrincipal = New TL_Diario
End Sub
Private Sub CreateTaskPanel()
   Dim Group As TaskPanelGroup
   Dim Item As TaskPanelGroupItem
   
   mvarPane.wndTaskPanel.SetImageList mvarPane.imlTaskPanelIcons
   
   Set Group = mvarPane.wndTaskPanel.Groups.Add(0, "Calendário")
   
   Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
   Set Item.Control = mvarPane.FraCalendario
   mvarPane.FraCalendario.BackColor = Item.BackColor
   Group.Expanded = True
   Group.Expandable = False
    
    
'   Set Group = mvarPane.wndTaskPanel.Groups.Add(0, "Busca específica")
'   Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
'   Set Item.Control = mvarPane.fraBuscaDetalhada
'   mvarPane.fraBuscaDetalhada.BackColor = Item.BackColor
'   Group.Expanded = False
    
    mvarPane.wndTaskPanel.Reposition
End Sub

Private Sub PopulaRsDatas()
   Dim Sql As String
   Sql = "Select Distinct Cast(Convert(varchar, DTDIARIO, 103) as DateTime) [DTDIARIO]"
   Sql = Sql & " From ODIARIO"
   Sql = Sql & " Where IDLOJA=" & mvarSys.Propriedades("IDLOJA")
   Call mvarSys.xDb.Abretabela(Sql, RsDatas)
End Sub
