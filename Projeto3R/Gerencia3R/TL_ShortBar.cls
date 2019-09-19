VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_ShortBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private WithEvents mvarMe  As FrmShortBar
Attribute mvarMe.VB_VarHelpID = -1

Private Tl000 As TL_000
Private TlFin001 As TL_FIN001
'Private TlFin002 As TL_FIN002
'Private TlFin003 As TL_FIN003

'Private TlOpe001 As TL_OPE001
'Private TlOpe002 As TL_OPE002
Dim TLPrincipal As Object

Private mvarItem  As Integer
Private mvarTL_CONS   As String
Private mvarSys      As Object
Private mvarMDI      As Object
Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
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
Public Property Let Item(ByVal vData As Integer)
   mvarItem = vData
End Property
Public Property Get Item() As Integer
   Item = mvarItem
End Property
Private Sub CreateTaskPanel()
    Dim Group  As TaskPanelGroup
    Dim Item   As TaskPanelGroupItem
    Dim i      As Integer
    Dim Sql    As String
    Dim sGrupo As String
    Dim MyRs   As Object
    
   With mvarMe.wndTaskPanel
      .HotTrackStyle = xtpTaskPanelHighlightItem
      .SingleSelection = True
      Call .SetImageList(mvarMe.imlTaskPanelIcons)
                  
      Set Group = .Groups.Add(0, "Conteudo")
      Set Item = Group.Items.Add(1, "Principal", xtpTaskItemTypeLink, 2)
      Item.Visible = False
      Group.Visible = False
      
      Set Group = .Groups.Add(1, "Financeiro")
      Set Item = Group.Items.Add(1001, "Vendas", xtpTaskItemTypeLink, 2)
      'Set Item = Group.Items.Add(1002, "Faturamento", xtpTaskItemTypeLink, 2)
      'Set Item = Group.Items.Add(1003, "Fluxo de Caixa", xtpTaskItemTypeLink, 2)
      'Set Item = Group.Items.Add(1004, "Comissões", xtpTaskItemTypeLink, 2)
            
      Set Group = .Groups.Add(2, "Operacional")
      Set Item = Group.Items.Add(2001, "Produção Mensal", xtpTaskItemTypeLink, 2)
      Set Item = Group.Items.Add(2002, "Produção Por Operadora", xtpTaskItemTypeLink, 2)
      'Set Item = Group.Items.Add(2003, "Produção Por Tratamento/Área", xtpTaskItemTypeLink, 2)
            
            
      'Set Group = .wndTaskPanel.Groups.Add(0, "Grupo3")
      'Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
      'Set Item.Control = .fraBuscaDetalhada
      '.fraBuscaDetalhada.BackColor = Item.BackColor
      'Group.Expanded = False
      If .Groups.Count = 1 Then
         .Groups(1).Expandable = False
      End If
      
      .Reposition
   End With
End Sub

Public Sub Show()
   '* Acoplar Form no MDI
   Load mvarMe
   Call AcoplarForm(mvarMe, 2, mvarSys)
   mvarMe.Show
   
   'mvarItem = IIf(mvarItem = 0, 1, mvarItem)
   If mvarMe.wndTaskPanel.FocusedItem Is Nothing Then
      Dim oItem As TaskPanelGroupItem
      Set oItem = mvarMe.wndTaskPanel.Find(mvarItem)
      If oItem Is Nothing Then
         If mvarMe.wndTaskPanel.Groups(1).Items.Count > 0 Then
            mvarMe.wndTaskPanel.Groups(1).Items(1).Selected = True
         End If
      Else
         oItem.Selected = True
      End If
   End If
   If Not mvarMe.wndTaskPanel.FocusedItem Is Nothing Then
      Call mvarMe_wndTaskPanelItemClick(mvarMe.wndTaskPanel.FocusedItem)
   End If
   Call mvarMe_Resize
End Sub
Public Sub Hide()
   mvarMe.Hide
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmShortBar
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_Activate()
   If GetTag(Me, "1Vez", 0) = 0 Then
      Call mvarMe_Resize
   End If
End Sub
Private Sub mvarMe_Load()
   Call CreateTaskPanel
End Sub
Private Sub mvarMe_Resize()
   On Error Resume Next
   With mvarMe
      .SccContato.Move 0, 0, .ScaleWidth, 420
      .SccContato2.Move 0, .SccContato.Height, .SccContato.Width, 285
   
      .Picture1.Move 0, .SccContato2.Top + .SccContato2.Height, .ScaleWidth, .ScaleHeight - .SccContato.Height
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
   Dim MyObj As Object
   
   If Not TLPrincipal Is Nothing Then Call TLPrincipal.Hide
   If Item.Group.Caption = "Conteudo" Then
      Select Case Item.Id
         Case 1 '* Principal
            If Tl000 Is Nothing Then Set Tl000 = New TL_000
            Set Tl000.Sys = mvarSys
            Call TLPrincipal.Show
      End Select
   
   ElseIf Item.Group.Caption = "Financeiro" Then
      Select Case Item.Id
         Case 1001 '* Vendas
            If TlFin001 Is Nothing Then Set TlFin001 = New TL_FIN001
            Set TlFin001.Sys = mvarSys
            Call TlFin001.Show
      End Select
   ElseIf Item.Group.Caption = "Operacional" Then
            If Tl000 Is Nothing Then Set Tl000 = New TL_000
            Set Tl000.Sys = mvarSys
            Call Tl000.Show
   Else
      If Tl000 Is Nothing Then Set Tl000 = New TL_000
      Set Tl000.Sys = mvarSys
      Call Tl000.Show
   End If
   If Not TLPrincipal Is Nothing Then
      Set TLPrincipal.Sys = mvarSys
      TLPrincipal.Show bReload
   End If
   
End Sub
Private Function DefineCad(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem) As String
   Select Case Item.Caption
      Case "Funcionário": DefineCad = "TL_CADRFUNCIONARIO"
  End Select
End Function
