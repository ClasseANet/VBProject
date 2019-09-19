VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_PaneCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe  As FrmPaneFinanc
Attribute mvarMe.VB_VarHelpID = -1
Private WithEvents mvarCmmdBars As XtremeCommandBars.CommandBars
Attribute mvarCmmdBars.VB_VarHelpID = -1
Dim TLPrincipal As Object

Dim MyCADFCCORRENTE As TL_CADFCCORRENTE
Dim MyMovCC As TL_MovCC

Private Enum TaskCC
   Mov = 1
   Det = 2
   His = 3
End Enum
Private Enum eMenuFin
   Mnu_Financ = 2400
   Mnu_ExportarMov
End Enum

Private mvarSys      As Object
Private mvarItemCad  As Integer
Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
    Set TLPrincipal.Sys = mvarSys
    Set mvarCmmdBars = mvarSys.MDI.CommandBars
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Property Get Pane() As Object
    Set Pane = mvarMe
End Property
Public Sub Show(Optional Reload As Boolean = False)
   Dim xItem As TaskPanelGroupItem
   If Reload Then Set Sys = Sys
   
   With mvarMe.wndTaskPanel
      Set xItem = .Groups(1).Items(Val(GetTag(mvarMe.wndTaskPanel, "Item0", 1)))
   End With
   xItem.Selected = True
   Call mvarMe_wndTaskPanelItemClick(xItem)
   
   '* Exibir Menu no MDI
   If Not mvarSys.MDI.CommandBars.FindControl(, eMenuFin.Mnu_Financ, False) Is Nothing Then
      mvarSys.MDI.CommandBars.FindControl(, eMenuFin.Mnu_Financ).Visible = True
   Else
      Call MontarMenu
   End If

   If Reload Then Call mvarMe_Resize
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
      Case "FINANC"
         Select Case Control.Id
            Case eMenuFin.Mnu_ExportarMov: Call ExportarMov
         End Select
   End Select
End Sub
Private Sub mvarMe_CommandBarsExecute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   If Control.Category = "Popup_IDCONTA" Then
      If TypeName(TLPrincipal) = "TL_MovCC" Then
         mvarSys.Propriedades("IDCONTA") = Control.Id
         TLPrincipal.IDCONTA = Control.Id
         TLPrincipal.PopulaTela
         mvarMe.SccConta2.Caption = "Conta " & Control.Caption
      ElseIf TypeName(TLPrincipal) = "TL_CADFCCORRENTE" Then
         mvarSys.Propriedades("IDCONTA") = Control.Id
         TLPrincipal.IDCONTA = Control.Id
         TLPrincipal.PopulaTela
         mvarMe.SccConta2.Caption = "Conta " & Control.Caption
      
      End If
   End If
End Sub
Private Sub mvarMe_Load()
   mvarMe.CommandBars.DeleteAll
   Call CreateTaskPanel
   Call MontaTela
   Call MontarMenu
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
Private Sub MontaTela()
   With mvarMe.CmbPeriodo
      .Clear
      .AddItem " TUDO ATÉ HOJE"
      .ItemData(.NewIndex) = 0
      .AddItem "Último mês"
      .ItemData(.NewIndex) = 30
      .AddItem "Últimos 03 meses"
      .ItemData(.NewIndex) = 90
      .AddItem "Últimos 06 meses"
      .ItemData(.NewIndex) = 180
      .AddItem "Últimos 12 meses"
      .ItemData(.NewIndex) = 360
      
      .ListIndex = 1
   End With
End Sub

Private Sub mvarMe_wndTaskPanelItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
   Screen.MousePointer = vbHourglass
   
   'TLPrincipal.Hide
   Select Case Item.Id
      Case TaskCC.Mov:   'MOVIMENTACAO
         If MyMovCC Is Nothing Then Set MyMovCC = New TL_MovCC
         Set MyMovCC.Pane = mvarMe
         Set TLPrincipal = MyMovCC
      
      Case TaskCC.Det: 'DETALHES
         If MyCADFCCORRENTE Is Nothing Then Set MyCADFCCORRENTE = New TL_CADFCCORRENTE
         Set TLPrincipal = MyCADFCCORRENTE

'      Case TaskCC.His: 'HISTORICO
'         If TypeName(TLPrincipal) <> "TL_MovCC" Then
'            Set TLPrincipal = Nothing
'            Set TLPrincipal = New TL_CCHistorico
'         End If
         
      Case Else
         Call ExibirAviso("Função não disponível nesta versão.")
         mvarMe.wndTaskPanel.Groups(1).Items(Val(GetTag(mvarMe.wndTaskPanel, "Item0", mvarMe.wndTaskPanel.FocusedItem.Index))).Selected = True
   End Select
   If Not TLPrincipal Is Nothing Then
      Set TLPrincipal.Sys = mvarSys
      Call TLPrincipal.Show
   End If
   Call SetTag(mvarMe.wndTaskPanel, "Item0", mvarMe.wndTaskPanel.FocusedItem.Index)
   Screen.MousePointer = vbDefault
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmPaneFinanc
   Set TLPrincipal = New TL_MovCC
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
         .Items.Add TaskCC.Det, "Detalhes    ", xtpTaskItemTypeLink, 2
         .Items.Add TaskCC.His, "Histórico   ", xtpTaskItemTypeLink, 2
         .Expandable = False
      End With
      
      Set Group = .Groups.Add(0, "Filtros")
      With Group
         Set Item = .Items.Add(0, "", xtpTaskItemTypeControl)
         Set Item.Control = mvarMe.FraFiltro
         mvarMe.FraFiltro.BackColor = Item.BackColor
         mvarMe.CmbPeriodo.BackColor = Item.BackColor
         .Expanded = True
      End With
      
      .Reposition
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
         Set oCommBar = .Add("Menu Financeiro", xtpBarTop)
         oCommBar.Closeable = False
         oCommBar.Customizable = False
         oCommBar.EnableDocking xtpFlagHideWrap
      Else
         Set oCommBar = .ActiveMenuBar
      End If
      
      Set oMenuMain = oCommBar.FindControl(, eMenuFin.Mnu_Financ)
      If Not oMenuMain Is Nothing Then
         If (oMenuMain.Category = "FINANC") Then
            oMenuMain.Visible = True
            GoTo Saida
         End If
      End If
      
      With oCommBar
         '************
         '* Menu Contato
         Set oMenuMain = AddButtonBar(.Controls, eMenuFin.Mnu_Financ, "Financeiro", False, xtpControlPopup, "FINANC")
         With oMenuMain.CommandBar
            AddButtonBar .Controls, eMenuFin.Mnu_ExportarMov, "Exportar Movimento", Category:="FINANC"
         End With
         'Call MontarContextMenu
      End With
      .RecalcLayout
   End With
GoTo Saida

TrataErro:
   'Resume Next
   MsgBox "Erro em Sub MontarMenu. Erro: " & Err.Number & "-" & Err.Description
Saida:
End Sub
Private Sub ExportarMov()
   Dim MyTela As TL_ExportMov
   
   Set MyTela = New TL_ExportMov
   Set MyTela.Sys = mvarSys
   MyTela.IDCONTA = TLPrincipal.IDCONTA
   MyTela.Show vbModal
End Sub

