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

Private Enum TaskCC
   Mov = 1
   Det = 2
   His = 3
End Enum

Private WithEvents mvarMe  As FrmPaneTarefa
Attribute mvarMe.VB_VarHelpID = -1
Dim TLTarefas As Object

Private mvarSys      As Object
Private mvarItemCad  As Integer
Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
    Set TLTarefas.Sys = mvarSys
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Property Get Pane() As Object
    Set Pane = mvarMe
End Property
Public Sub Show()
   If mvarMe.wndTaskPanel.FocusedItem Is Nothing Then
      mvarMe.wndTaskPanel.Groups(1).Items(1).Selected = True
   End If
   Call mvarMe_wndTaskPanelItemClick(mvarMe.wndTaskPanel.FocusedItem)
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub

Private Sub mvarMe_CommandBarsExecute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   If Control.Category = "Popup_IDCONTA" Then
      If TypeName(TLTarefas) = "TL_MovCC" Then
         mvarSys.Propriedades("IDCONTA") = Control.Id
         TLTarefas.IDCONTA = Control.Id
         TLTarefas.PopulaTela
         mvarMe.SccConta2.Caption = "Conta " & Control.Caption
      End If
   End If
End Sub
Private Sub mvarMe_Load()
   mvarMe.CommandBars.DeleteAll
   Call CreateTaskPanel
End Sub
Private Sub mvarMe_Resize()
   With mvarMe
      .SccConta.Move 0, 0, .ScaleWidth, 420
      .SccConta2.Move 0, .SccConta.Height, .SccConta.Width, 285
   
      .Picture1.Move 0, .SccConta2.Top + .SccConta2.Height, .ScaleWidth, .ScaleHeight - .SccConta.Height
      .wndTaskPanel.Move 0, 0, .Picture1.ScaleWidth, .Picture1.ScaleHeight
   End With
End Sub

Private Sub mvarMe_SccConta2MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim Popup     As CommandBar
   Dim oCmd    As CommandBarControl
   Dim TbCC    As Object
   
   Set TbCC = CriarObjeto("BANCO_3R.TB_FCCORRENTE")
   Set TbCC.xDb = mvarSys.xDb
   If TbCC.Pesquisar(Ch_IDLOJA:=mvarSys.Propriedades("IDLOJA")) Then
      Set Popup = mvarMe.CommandBars.Add("Popup", xtpBarPopup)
      With Popup
         .Controls.DeleteAll
         While Not TbCC.Rs.EOF
            TbCC.Popula
            Set oCmd = .Controls.Add(XTPControlType.xtpControlButton, TbCC.IDCONTA, TbCC.DSCCONTA)
            oCmd.Category = "Popup_IDCONTA"
            oCmd.Checked = (Val(mvarSys.Propriedades("IDCONTA")) = TbCC.IDCONTA)
            TbCC.Rs.MoveNext
         Wend
        .ShowPopup
        .Controls.DeleteAll
      End With
   End If
End Sub

Private Sub mvarMe_wndTaskPanelFocusedItemChanged()
   Dim nGrp As TaskPanelGroup
   Dim nItm As TaskPanelGroupItem
   
   With mvarMe.wndTaskPanel
      For Each nGrp In .Groups
         For Each nItm In nGrp.Items
            nItm.IconIndex = 2
         Next
      Next
      If Not .FocusedItem.IsGroup Then
         .FocusedItem.IconIndex = 1
      End If
   End With
End Sub
Private Sub mvarMe_wndTaskPanelItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
   Screen.MousePointer = vbHourglass
   
   Select Case Item.Id
      
      Case TaskCC.Mov:   'MOVIMENTACAO
      
         If TypeName(TLTarefas) <> "TL_Tarefas" Then
            Set TLTarefas = Nothing
            Set TLTarefas = New TL_Tarefas
         End If
'      Case TaskCC.Det: 'DETALHES
'         If TypeName(TLTarefas) <> "TL_CADCCORRENTE" Then
'            Set TLTarefas = Nothing
'            Set TLTarefas = New TL_CADFCCORRENTE
'         End If
'      Case TaskCC.His: 'HISTORICO
'         If TypeName(TLTarefas) <> "TL_MovCC" Then
'            Set TLTarefas = Nothing
'            Set TLTarefas = New TL_CCHistorico
'         End If
         
      Case Else
         Call ExibirAviso("Função não disponível nesta versão.")
         mvarMe.wndTaskPanel.Groups(1).Items(TaskCC.Mov).Selected = True
   End Select
   
   If Not TLTarefas Is Nothing Then
      Set TLTarefas.Sys = mvarSys
      'Set TLTarefas.MDI = mvarMDI
      Call TLTarefas.Show
   End If
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmPaneTarefa
   Set TLTarefas = New TL_Tarefas
End Sub
Private Sub CreateTaskPanel()
   Dim Group As TaskPanelGroup
   Dim Item   As TaskPanelGroupItem
    
   With mvarMe.wndTaskPanel
      .HotTrackStyle = xtpTaskPanelHighlightItem
      .SingleSelection = True
      Call .SetImageList(mvarMe.imlTaskPanelIcons)
    
      Set Group = .Groups.Add(0, "Detalhes da Conta")
      With Group
         .Items.Add TaskCC.Mov, "Movimentação", xtpTaskItemTypeLink, 1
         .Items.Add TaskCC.Det, "Cadastro    ", xtpTaskItemTypeLink, 2
         .Items.Add TaskCC.His, "Histórico   ", xtpTaskItemTypeLink, 2
         .Expandable = False
      End With
      .Reposition
   End With
End Sub
