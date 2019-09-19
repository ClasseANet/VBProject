VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_FIN001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe  As FrmFIN001
Attribute mvarMe.VB_VarHelpID = -1
Private mvarSys      As Object
Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Sub Show(Optional Reload As Boolean = False)
   mvarMe.Show
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_Load()
   With mvarMe
      .CommandBars.DeleteAll
   
      With .DpiCalendario
         .Move 0, 0, .Width, 2400
         .ShowNonMonthDays = False
         .FirstDayOfWeek = 1
         .ShowNoneButton = False
         .AskDayMetrics = True
       End With
      .FraCalendario.Height = 2400
      
   End With
   Call CreateTaskPanel
End Sub
Private Sub mvarMe_Resize()
   With mvarMe
      .SccConta.Move 0, 0, .ScaleWidth, 420
      .SccConta2.Move 0, .SccConta.Height, .SccConta.Width, 285
   
      .Picture1.Move 0, .SccConta2.Top + .SccConta2.Height, .ScaleWidth, .ScaleHeight - .SccConta.Height
      .wndTaskPanel.Move 0, 0, .Picture1.ScaleWidth, .Picture1.ScaleHeight
      
      mvarMe.DpiCalendario.ShowNoneButton = False
      
   End With
End Sub

Private Sub mvarMe_wndTaskPanelItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
   Dim nRow As Long
   Screen.MousePointer = vbHourglass
   
   
   
   Select Case Item.Id
      Case 1:       'ADD CONTATO
      Case 2:
         'Call TLConsulta.EditarObjeto
         
      Case 3:       'DEL CONTATO
      Case 4:       '
      Case 5:       '
      Case 6:       '
      
   End Select
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub Class_Initialize()
'   Set TLConsulta = New TL_Consulta3R
End Sub
Private Sub CreateTaskPanel()
'   Dim Group As TaskPanelGroup
'   Dim Item As TaskPanelGroupItem
'
'   mvarMe.wndTaskPanel.SetImageList mvarMe.imlTaskPanelIcons
'
'   Set Group = mvarMe.wndTaskPanel.Groups.Add(0, "Operações")
'   Group.Items.Add TaskItem.NovaTarefa, "Adicionar nova tarefa", xtpTaskItemTypeLink, 4
'   Group.Items.Add TaskItem.EditTarefa, "Editar tarefa selecionada", xtpTaskItemTypeLink, 5
'   Set Item = Group.Items.Add(TaskItem.ExclTarefa, "Excluir tarefa selecionada", xtpTaskItemTypeLink, 6)
'   Item.Enabled = False
'   Group.Expandable = False
'
'   Set Group = mvarMe.wndTaskPanel.Groups.Add(0, "Calendário")
'   Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
'   Set Item.Control = mvarMe.FraCalendario
'   mvarMe.FraCalendario.BackColor = Item.BackColor
'   Group.Expanded = True
'   Group.Expandable = False
'
'
''   Set Group = mvarMe.wndTaskPanel.Groups.Add(0, "Busca específica")
''   Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
''   Set Item.Control = mvarMe.fraBuscaDetalhada
''   mvarMe.fraBuscaDetalhada.BackColor = Item.BackColor
''   Group.Expanded = False
'
'    mvarMe.wndTaskPanel.Reposition
End Sub

