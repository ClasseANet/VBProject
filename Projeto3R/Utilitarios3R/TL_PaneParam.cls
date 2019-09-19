VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_PaneParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Enum Task
   Aplicacao = 1 'Ambiente = 2
   Interface
   Com
   Fin
   Log
   Geral
End Enum

Private WithEvents mvarMe  As FrmPaneParam
Attribute mvarMe.VB_VarHelpID = -1
Dim TLDefault As Object

Private mvarsys      As Object
Private mvarMDI      As Object
Public Property Set Sys(ByVal vData As Object)
    Set mvarsys = vData
    Set TLDefault.Sys = mvarsys
End Property
Public Property Get Sys() As Object
    Set Sys = mvarsys
End Property
Public Property Get Pane() As Object
    Set Pane = mvarMe
End Property
Public Property Set MDI(ByVal vData As Object)
    Set mvarMDI = vData
End Property
Public Property Get MDI() As Object
    Set MDI = mvarMDI
End Property
Public Sub Show()
   '* Carregar e Acoplar Form no MDI
   Load mvarMe
   Call AcoplarForm(mvarMe, 1, mvarsys, pMDI:=mvarsys.ObjAux)
   
   If mvarMe.wndTaskPanel.FocusedItem Is Nothing Then
      mvarMe.wndTaskPanel.Groups(1).Items(Task.Aplicacao).Selected = True
   End If
   Call mvarMe_wndTaskPanelItemClick(mvarMe.wndTaskPanel.FocusedItem)
End Sub
Private Sub mvarMe_Load()
   Call CreateTaskPanel
End Sub
Private Sub mvarMe_Resize()
   With mvarMe
      .SccTit1.Move 0, 0, .ScaleWidth, 420
      .SccTit2.Move 0, .SccTit1.Height, .SccTit1.Width, 285
   
      .Picture1.Move 0, .SccTit2.Top + .SccTit2.Height, .ScaleWidth, .ScaleHeight - .SccTit1.Height
      .wndTaskPanel.Move 0, 0, .Picture1.ScaleWidth, .Picture1.ScaleHeight
   End With
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
      Case Task.Aplicacao
         If TypeName(TLDefault) <> "TL_ParamAplicacao" Then
            Set TLDefault = Nothing
            Set TLDefault = New TL_ParamAplicacao
         End If
      
      Case Task.Com
         If TypeName(TLDefault) <> "TL_ParamCom" Then
            Set TLDefault = Nothing
            Set TLDefault = New TL_ParamCom
         End If
      
      Case Task.Fin
         If TypeName(TLDefault) <> "TL_ParamFin" Then
            Set TLDefault = Nothing
            Set TLDefault = New TL_ParamFin
         End If
      
      Case Task.Geral
         If TypeName(TLDefault) <> "TL_ParamGeral" Then
            Set TLDefault = Nothing
            Set TLDefault = New TL_ParamGeral
         End If
      
      Case Else
         'Set TLDefault = Nothing
         'Call ExibirAviso(""Função não disponível nesta versão."")
         'mvarMe.wndTaskPanel.Groups(1).Items(Task.Com).Selected = True
         
         Call ExibirAviso("Função não disponível nesta versão.")
         mvarMe.wndTaskPanel.Groups(1).Items(Task.Com).Selected = True
         
   End Select
   
   If Not TLDefault Is Nothing Then
      Set TLDefault.Sys = mvarsys
      Set TLDefault.MDI = mvarMDI
      Call TLDefault.Show
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmPaneParam
   Set TLDefault = New TL_ParamCom
End Sub
Private Sub CreateTaskPanel()
   Dim Group As TaskPanelGroup
   Dim Item   As TaskPanelGroupItem
    
   With mvarMe.wndTaskPanel
      .HotTrackStyle = xtpTaskPanelHighlightItem
      .SingleSelection = True
      Call .SetImageList(mvarMe.imlTaskPanelIcons)
    
      Set Group = .Groups.Add(0, "Configurações")
      With Group
         .Items.Add Task.Aplicacao, "Aplicação", xtpTaskItemTypeLink, 1
         '.Items.Add Task.Interface, "Interface de Usuário", xtpTaskItemTypeLink, 1
         .Items.Add Task.Com, "Comunicação", xtpTaskItemTypeLink, 1
         .Items.Add Task.Fin, "Financeiro", xtpTaskItemTypeLink, 1
         '.Items.Add Task.Log, "Log de Eventos", xtpTaskItemTypeLink, 1
         .Items.Add Task.Geral, "Geral", xtpTaskItemTypeLink, 1
         .Expandable = False
      End With
      .Reposition
   End With
End Sub


