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
Private mvarTLResult As TL_FIN001

Private mvarItemCons  As Integer
Private mvarTL_CONS   As String
Private mvarSys      As Object
Private mvarMDI      As Object
Public Property Get TLResult() As TL_FIN001
   Set TLResult = mvarTLResult
End Property
Public Property Let TL_CAD(ByVal vData As String)
   mvarTL_CONS = vData
End Property
Public Property Get TL_CAD() As String
   TL_CAD = mvarTL_CONS
End Property
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
Public Property Let ItemCons(ByVal vData As Integer)
   mvarItemCons = vData
End Property
Public Property Get ItemCons() As Integer
   ItemCons = mvarItemCons
End Property
Private Sub CreateTaskPanel()
    Dim Group  As TaskPanelGroup
    Dim Item   As TaskPanelGroupItem
    Dim i      As Integer
    Dim SQL    As String
    Dim sGrupo As String
    Dim MyRs   As Object
    
   With mvarMe.wndTaskPanel
      .HotTrackStyle = xtpTaskPanelHighlightItem
      .SingleSelection = True
      Call .SetImageList(mvarMe.imlTaskPanelIcons)
      
      Set Group = .Groups.Add(1, "Financeiro")
      Set Item = Group.Items.Add(1001, "Vendas", xtpTaskItemTypeLink, 2)
      Set Item = Group.Items.Add(1002, "Faturamento", xtpTaskItemTypeLink, 2)
      Set Item = Group.Items.Add(1003, "Fluxo de Caixa", xtpTaskItemTypeLink, 2)
      Set Item = Group.Items.Add(1004, "Comissões", xtpTaskItemTypeLink, 2)
            
      Set Group = .Groups.Add(2, "Operacional")
      Set Item = Group.Items.Add(2001, "Produção Mensal", xtpTaskItemTypeLink, 2)
      Set Item = Group.Items.Add(2002, "Produção Por Operadora", xtpTaskItemTypeLink, 2)
      Set Item = Group.Items.Add(2003, "Produção Por Tratamento/Área", xtpTaskItemTypeLink, 2)
            
            
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
   '* Carregar e Acoplar Form no MDI
   Load mvarMe
   Call AcoplarForm(mvarMe, 1, mvarSys, pMDI:=mvarMDI) ' mvarSys.ObjAux)
   
   mvarItemCons = IIf(mvarItemCons = 0, 1, mvarItemCons)
   If mvarMe.wndTaskPanel.FocusedItem Is Nothing Then
      Dim oItem As TaskPanelGroupItem
      Set oItem = mvarMe.wndTaskPanel.Find(mvarItemCons)
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
   
Exit Sub
   
   Load mvarMe
   If mvarSys.MDI.DockingPaneManager.Panes(1).Handle <> mvarMe.hwnd Then
      SetMDI mvarMe.hwnd, mvarSys.MDI.hwnd
      mvarSys.MDI.DockingPaneManager.Panes(1).Handle = mvarMe.hwnd
   End If
   
   mvarItemCons = 19
   
   Dim xItem As XtremeTaskPanel.ITaskPanelGroupItem
   If mvarItemCons > 0 Then
      Set xItem = mvarMe.wndTaskPanel.Find(mvarItemCons)
      If Not xItem Is Nothing Then
         Call mvarMe_wndTaskPanelItemClick(xItem)
      End If
      'Call mvarMe_wndTaskPanelItemClick(xItem)
   End If
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmShortBar
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Set mvarMe = Nothing
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
   Dim SQL     As String
   Dim sScript As String
   Dim sTl_Cad As String
   Dim sTag    As String
   Dim sIDMODU As String
         
   SQL = Trim(GetTag(Item, "CONSULTA", ""))
   sScript = Trim(GetTag(Item, "VBSCRIPT", ""))
   sTl_Cad = Trim(GetTag(Item, "TL_CAD", ""))
   sTag = Trim(Decrypt2(GetTag(Item, "TAGCAMPOS", "")))
   sIDMODU = Trim(GetTag(Item, "IDMODU", ""))
   
   SQL = IIf(SQL = "", sScript, SQL)
   mvarTL_CONS = ""
   If UCase(Mid(SQL, 1, Len("Select "))) = "SELECT " Then
      If sTl_Cad = "" Then
         sTl_Cad = DefineCad(Item)
      End If
      mvarTL_CONS = sTl_Cad
      
      Set mvarTLResult = Nothing
      Set mvarTLResult = New TL_FIN001
      
      With mvarTLResult
         Set .Sys = mvarSys
         .Show
      End With
   ElseIf SQL <> "" Then
      Set mvarTLResult = Nothing
      Set mvarTLResult = CriarObjeto(SQL)
      If Not mvarTLResult Is Nothing Then
         Set mvarTLResult.Sys = mvarSys
         '* Acoplar Form no MDI
         'Call AcoplarForm(mvarTLResult.frm, 2, mvarSys, True, mvarMDI)
         mvarTLResult.Show
      End If
   End If
End Sub
Private Function DefineCad(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem) As String
   Select Case Item.Caption
      Case "Funcionário": DefineCad = "TL_CADRFUNCIONARIO"
  End Select
End Function
