VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_ShortBar0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private WithEvents mvarMe  As FrmShortBar
Attribute mvarMe.VB_VarHelpID = -1
Private mvarTLCadastro As TL_Cadastro

Private mvarItemCad  As Integer
Private mvarTL_CAD   As String
Private mvarSys      As Object
Private mvarMDI      As Object
Public Property Get TLCadastro() As TL_Cadastro
   Set TLCadastro = mvarTLCadastro
End Property
Public Property Let TL_CAD(ByVal vData As String)
   mvarTL_CAD = vData
End Property
Public Property Get TL_CAD() As String
   TL_CAD = mvarTL_CAD
End Property
Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
    Set mvarTLCadastro.Sys = mvarSys
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
Public Property Let ItemCad(ByVal vData As Integer)
   mvarItemCad = vData
End Property
Public Property Get ItemCad() As Integer
   ItemCad = mvarItemCad
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
            
      SQL = "Select C.ID, M.IDMODU, C.GRUPO, M.DSCMODU, C.CONSULTA"
      SQL = SQL & ", M.VBSCRIPT, C.TL_CAD, C.TAGCAMPOS, IsNull(U.IDUSU, 'DIO') [IDUSU]"
      SQL = SQL & " From  GCAD C Join MODULO M On C.ID=M.ID"
      SQL = SQL & " Join MODULO_SISTEMA S On S.ID=M.ID"
      SQL = SQL & " Left Join GRPACESSO G On M.ID=G.ID"
      SQL = SQL & " Left Join USU_GRUPOS U On G.IDGRUPO=G.IDGRUPO"
      SQL = SQL & " Where M.SITMODU = " & SqlStr("S")
      SQL = SQL & " And S.CODSIS = " & SqlStr(mvarSys.CODSIS)
      SQL = SQL & " AND IsNull(U.IDUSU, 'DIO') =" & SqlStr(mvarSys.IDUSU)
      SQL = SQL & " Order By C.GRUPO, M.DSCMODU"
      
      
      sGrupo = "Principal"
      Set Group = .Groups.Add(0, sGrupo)
      
      If mvarSys.XDb.AbreTabela(SQL, MyRs) Then
         While Not MyRs.EOF
            If Trim(MyRs("GRUPO") & "") <> "" And sGrupo <> Trim(MyRs("GRUPO") & "") Then
               i = i + 1
               Set Group = .Groups.Add(i, sGrupo)
               .Groups(1).Expandable = True
            End If
            sGrupo = IIf(Trim(MyRs("GRUPO") & "") = "", "Principal", Trim(MyRs("GRUPO") & ""))
            Set Item = Group.Items.Add(MyRs("ID"), MyRs("DSCMODU"), xtpTaskItemTypeLink, 2)
            Call SetTag(Item, "IDMODU", MyRs("IDMODU") & "")
            Call SetTag(Item, "CONSULTA", MyRs("CONSULTA") & "")
            Call SetTag(Item, "VBSCRIPT", MyRs("VBSCRIPT") & "")
            Call SetTag(Item, "TL_CAD", MyRs("TL_CAD") & "")
            Call SetTag(Item, "TAGCAMPOS", Encrypt2(MyRs("TAGCAMPOS") & ""))
            MyRs.MoveNext
         Wend
      End If
      
      
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
   
   mvarItemCad = IIf(mvarItemCad = 0, 1, mvarItemCad)
   If mvarMe.wndTaskPanel.FocusedItem Is Nothing Then
      Dim oItem As TaskPanelGroupItem
      Set oItem = mvarMe.wndTaskPanel.Find(mvarItemCad)
      If oItem Is Nothing Then
         mvarMe.wndTaskPanel.Groups(1).Items(1).Selected = True
      Else
         oItem.Selected = True
      End If
   End If
   Call mvarMe_wndTaskPanelItemClick(mvarMe.wndTaskPanel.FocusedItem)
   
Exit Sub
   
   Load mvarMe
   If mvarSys.MDI.DockingPaneManager.Panes(1).Handle <> mvarMe.hwnd Then
      SetMDI mvarMe.hwnd, mvarSys.MDI.hwnd
      mvarSys.MDI.DockingPaneManager.Panes(1).Handle = mvarMe.hwnd
   End If
   
   mvarItemCad = 19
   
   Dim xItem As XtremeTaskPanel.ITaskPanelGroupItem
   If mvarItemCad > 0 Then
      Set xItem = mvarMe.wndTaskPanel.Find(mvarItemCad)
      If Not xItem Is Nothing Then
         Call mvarMe_wndTaskPanelItemClick(xItem)
      End If
      'Call mvarMe_wndTaskPanelItemClick(xItem)
   End If
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmShortBar
   Set mvarTLCadastro = New TL_Cadastro
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
   mvarTL_CAD = ""
   If UCase(Mid(SQL, 1, Len("Select "))) = "SELECT " Then
      If sTl_Cad = "" Then
         sTl_Cad = DefineCad(Item)
      End If
      mvarTL_CAD = sTl_Cad
      
      Set mvarTLCadastro = Nothing
      Set mvarTLCadastro = New TL_Cadastro
      
      With mvarTLCadastro
         Set .Sys = mvarSys
         .Consulta = SQL
         .Titulo = Item.Caption
         .TL_CAD = sTl_Cad
         .TAGCAMPOS = sTag
         .IDMODU = sIDMODU
         
         Set .MDI = mvarMDI
         'Set mvarSys.ObjAux = mvarMDI
         .Show
      End With
   Else
      Set mvarTLCadastro = Nothing
      Set mvarTLCadastro = CriarObjeto(SQL)
      If Not mvarTLCadastro Is Nothing Then
         Set mvarTLCadastro.Sys = mvarSys
         '* Acoplar Form no MDI
         'Call AcoplarForm(mvarTLCadastro.frm, 2, mvarSys, True, mvarMDI)
         mvarTLCadastro.Show
      End If
   End If
End Sub
Private Function DefineCad(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem) As String
   Select Case Item.Caption
      Case "Funcionário": DefineCad = "TL_CADRFUNCIONARIO"
  End Select
End Function
