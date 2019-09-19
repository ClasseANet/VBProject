VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_PaneSupri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Enum TaskSupri
   Prd = 1
   Mov
   Inv
   AddProd
   EdtProd
   DelProd
   EstEnt
   EstSai
   EstCont
End Enum
Private Enum eMenuSup
   Mnu_Supri = 2500
   Mnu_EstEnt
   Mnu_EstSai
   Mnu_EstCont
End Enum


Private WithEvents mvarMe  As FrmPaneSupri
Attribute mvarMe.VB_VarHelpID = -1
Private WithEvents mvarCmmdBars As XtremeCommandBars.CommandBars
Attribute mvarCmmdBars.VB_VarHelpID = -1
Dim TLPrincipal As Object
Dim bReload As Boolean

Dim MyTL_MovEst As TL_MovEst
Dim MyTL_Prod As TL_Produtos
Dim MyTL_Inv As TL_Inventario

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
   bReload = Reload
   If Reload Then Set Sys = Sys
   
   With mvarMe.wndTaskPanel
      Set xItem = .Groups(1).Items(Val(GetTag(mvarMe.wndTaskPanel, "Item0", 2)))
   End With
   xItem.Selected = True
   Call mvarMe_wndTaskPanelItemClick(xItem)
   
   '* Exibir Menu no MDI
   If Not mvarSys.MDI.CommandBars.FindControl(, eMenuSup.Mnu_Supri, False) Is Nothing Then
      mvarSys.MDI.CommandBars.FindControl(, eMenuSup.Mnu_Supri).Visible = True
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
Private Sub mvarMe_CommandBarsExecute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   If Control.Category = "Popup_IDPROD" Then
      If TypeName(TLPrincipal) = "TL_MovEst" Then
         mvarSys.Propriedades("IDPROD") = Control.Id
         TLPrincipal.IDPROD = Control.Id
         TLPrincipal.PopulaTela
         mvarMe.SccTit2.Caption = "Produto: " & Control.Caption
      ElseIf TypeName(TLPrincipal) = "TL_CADFCCORRENTE" Then
         mvarSys.Propriedades("IDCONTA") = Control.Id
         TLPrincipal.IDCONTA = Control.Id
         TLPrincipal.PopulaTela
         mvarMe.SccTit2.Caption = "Produto: " & Control.Caption
      
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
      .SccTit1.Move 0, 0, .ScaleWidth, 420
      .SccTit2.Move 0, .SccTit1.Height, .SccTit1.Width, 285
   
      .Picture1.Move 0, .SccTit2.Top + .SccTit2.Height, .ScaleWidth, .ScaleHeight - .SccTit1.Height
      .wndTaskPanel.Move 0, 0, .Picture1.ScaleWidth, .Picture1.ScaleHeight
   End With
End Sub

Private Sub mvarMe_Scctit2MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Exit Sub
   
   Dim Popup   As CommandBar
   Dim oCmd    As CommandBarControl
   Dim TbProd  As Object
   Dim Sql As String
   
   Set TbProd = CriarObjeto("BANCO_3R.TB_SPRODUTO")
   Set TbProd.xDb = mvarSys.xDb
   Sql = " ESERVICO=0"
   If TbProd.Pesquisar(Ch_IDLOJA:=mvarSys.IDLOJA, Ch_OrderBy:="NMPROD", Ch_Where:=Sql) Then
      Set Popup = mvarMe.CommandBars.Add("Popup", xtpBarPopup)
      With Popup
         .Controls.DeleteAll
         While Not TbProd.RS.EOF
            TbProd.Popula
            Set oCmd = .Controls.Add(XTPControlType.xtpControlButton, TbProd.IDPROD, TbProd.NMPROD)
            oCmd.Category = "Popup_IDPROD"
            oCmd.Checked = (Val(mvarSys.Propriedades("IDPROD")) = TbProd.IDPROD)
            TbProd.RS.MoveNext
         Wend
        .ShowPopup
        .Controls.DeleteAll
      End With
   End If
End Sub

Private Sub mvarMe_wndTaskPanelFocusedItemChanged()
   Dim nGrp As TaskPanelGroup
   Dim nItm As TaskPanelGroupItem
   Dim i As Integer
      
   With mvarMe.wndTaskPanel
      If .FocusedItem.Id > 3 Then Exit Sub
      For i = 1 To 3
         .Groups(1).Items(i).IconIndex = 2
      Next
      
      If Not .FocusedItem.IsGroup Then
         .FocusedItem.IconIndex = 1
      End If
   End With
End Sub
Private Sub mvarMe_wndTaskPanelItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
   If Item.Id > 3 Then Exit Sub
   On Error Resume Next
   Screen.MousePointer = vbHourglass
   
   Call TLPrincipal.Hide
   mvarMe.wndTaskPanel.Groups(1).Visible = True
   mvarMe.wndTaskPanel.Groups(2).Visible = False
   mvarMe.wndTaskPanel.Groups(3).Visible = False
   mvarMe.wndTaskPanel.Groups(4).Visible = False
   mvarMe.wndTaskPanel.Visible = False
   Select Case Item.Id
      Case TaskSupri.Prd:   'Produtos
         mvarMe.wndTaskPanel.Groups(2).Visible = True
         
         If MyTL_Prod Is Nothing Then
            Set MyTL_Prod = New TL_Produtos
            Set MyTL_Prod.Pane = mvarMe
         End If
         Set TLPrincipal = MyTL_Prod
      
      Case TaskSupri.Mov:   'MOVIMENTACAO
         mvarMe.wndTaskPanel.Groups(3).Visible = True
         
         If MyTL_MovEst Is Nothing Then Set MyTL_MovEst = New TL_MovEst
         Set MyTL_MovEst.Pane = mvarMe
         Set TLPrincipal = MyTL_MovEst
      Case TaskSupri.Inv: 'INVENTARIO
         mvarMe.wndTaskPanel.Groups(4).Visible = True
         If MyTL_Inv Is Nothing Then
            Set MyTL_Inv = New TL_Inventario
            Set MyTL_Inv.Pane = mvarMe
         End If
         Set TLPrincipal = MyTL_Inv

   End Select
   mvarMe.wndTaskPanel.Visible = True
   If Not TLPrincipal Is Nothing Then
      Set TLPrincipal.Sys = mvarSys
      TLPrincipal.Show bReload
   End If
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmPaneSupri
   Set TLPrincipal = New TL_MovEst
End Sub
Private Sub CreateTaskPanel()
   Dim Group As TaskPanelGroup
   Dim Item   As TaskPanelGroupItem
    
   With mvarMe.wndTaskPanel
      .HotTrackStyle = xtpTaskPanelHighlightItem
      .SingleSelection = True
      Call .SetImageList(mvarMe.imlTaskPanelIcons)
        
      Set Group = .Groups.Add(0, "Detalhes de Estoque")
      With Group
         .Items.Add TaskSupri.Prd, "Produtos/Serviços", xtpTaskItemTypeLink, 1
         .Items.Add TaskSupri.Mov, "Movimentação", xtpTaskItemTypeLink, 2
         .Items.Add TaskSupri.Inv, "Inventário  ", xtpTaskItemTypeLink, 2
         .Expandable = False
      End With
      
      Set Group = .Groups.Add(0, "Cadastro")
      With Group
         .Items.Add TaskSupri.AddProd, "Adicionar novo produto/serviço", xtpTaskItemTypeLink, 4
         .Items.Add TaskSupri.EdtProd, "Editar produto/serviço", xtpTaskItemTypeLink, 1
         .Items.Add TaskSupri.DelProd, "Excluir produto/serviço", xtpTaskItemTypeLink, 3
         .Expandable = False
         .Visible = False
      End With
      
      Set Group = .Groups.Add(0, "Filtros")
      With Group
         Set Item = .Items.Add(0, "", xtpTaskItemTypeControl)
         Set Item.Control = mvarMe.FraFiltro
         mvarMe.FraFiltro.BackColor = Item.BackColor
         mvarMe.CmbPeriodo.BackColor = Item.BackColor
         .Expandable = False
      End With
      
      Set Group = .Groups.Add(0, "Movimentos")
      With Group
         .Items.Add TaskSupri.EstEnt, "Entrada de Material", xtpTaskItemTypeLink, 5
         .Items.Add TaskSupri.EstSai, "Saída de Material", xtpTaskItemTypeLink, 6
         .Items.Add TaskSupri.EstCont, "Contagem de Estoque", xtpTaskItemTypeLink, 7
         .Expandable = False
         .Visible = False
      End With
      .Reposition
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
         Set oCommBar = .Add("Menu Suprimento", xtpBarTop)
         oCommBar.Closeable = False
         oCommBar.Customizable = False
         oCommBar.EnableDocking xtpFlagHideWrap
      Else
         Set oCommBar = .ActiveMenuBar
      End If
      
      Set oMenuMain = oCommBar.FindControl(, eMenuSup.Mnu_Supri)
      If Not oMenuMain Is Nothing Then
         If (oMenuMain.Category = "SUPRI") Then
            oMenuMain.Visible = True
            GoTo Saida
         End If
      End If
      
      With oCommBar
         '************
         '* Menu Contato
         Set oMenuMain = AddButtonBar(.Controls, eMenuSup.Mnu_Supri, "Suprimento", False, xtpControlPopup, "SUPRI")
         With oMenuMain.CommandBar
            AddButtonBar .Controls, eMenuSup.Mnu_EstEnt, "Entrada de Material", Category:="FINANC"
            AddButtonBar .Controls, eMenuSup.Mnu_EstSai, "Saída de Material", Category:="FINANC"
            AddButtonBar .Controls, eMenuSup.Mnu_EstCont, "Contagem de Estoque", Category:="FINANC", BeginGroup:=True
         End With
         'Call MontarContextMenu
      End With
      .RecalcLayout
   End With
GoTo Saida

TrataErro:
   'Resume Next
   MsgBox "Erro em MontarMenu. Erro: " & Err.Number & "-" & Err.Description
Saida:
End Sub

