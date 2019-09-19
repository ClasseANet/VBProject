VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_PANE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Const TaskCC_Mov = 1
Const TaskCC_Det = 2
Const TaskCC_His = 3

Private WithEvents mvarMe  As FrmPaneFinanc
Attribute mvarMe.VB_VarHelpID = -1
Dim TLFinanceiro As Object

Private mvarSys      As Object
Private mvarItemCad  As Integer
Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Private Sub CreateTaskPanel()
   Dim Group As TaskPanelGroup
   Dim Item   As TaskPanelGroupItem
    
   With mvarPane.wndTaskPanel
      .HotTrackStyle = xtpTaskPanelHighlightItem
      .SingleSelection = True
      Call .SetImageList(mvarPane.imlTaskPanelIcons)
    
      Set Group = .Groups.Add(0, "O que você deseja?")
      With Group
         .Items.Add TaskCC_Mov, "Movimentação", xtpTaskItemTypeLink, 17
         .Items.Add TaskCC_Det, "Detalhes    ", xtpTaskItemTypeLink, 18
         .Items.Add TaskCC_His, "Histórico   ", xtpTaskItemTypeLink, 19
         .Expandable = False
      End With
      .Reposition
   End With
End Sub
Private Sub mvarPane_Load()
   Call CreateTaskPanel
End Sub
Private Sub mvarPane_Resize()
   With mvarPane
      .SccConta.Move 0, 0, .ScaleWidth, 420
      .SccConta2.Move 0, .SccConta.Height, .SccConta.Width, 285
   
      .Picture1.Move 0, .SccConta2.Top + .SccConta2.Height, .ScaleWidth, .ScaleHeight - .SccConta.Height
      .wndTaskPanel.Move 0, 0, .Picture1.ScaleWidth, .Picture1.ScaleHeight
   End With
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmPaneFinanc
   Set TLFinanceiro = New TL_MovCC
End Sub
