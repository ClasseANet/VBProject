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
Option Explicit
Private WithEvents mvarMe  As FrmShortBar
Attribute mvarMe.VB_VarHelpID = -1
Private mvarTLCadastro As TL_Cadastro

Private mvarItemCad  As Integer
Private mvarTL_CAD   As String
Private mvarSys      As Object
Private mvarMDI      As Object
Public Property Get Frm() As Object
   Set Frm = mvarMe
End Property
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
            
      sGrupo = "Principal"
      Set Group = .Groups.Add(0, sGrupo)
      
      SQL = ""
      SQL = SQL & "Select *" & vbNewLine
      SQL = SQL & " From MODULO" & vbNewLine
      SQL = SQL & " Where VBSCRIPT like '%.NG_CADASTRO.%'" & vbNewLine
      SQL = SQL & " And SITMODU='S'" & vbNewLine
      SQL = SQL & " And MENUDEFAULT = 'S'" & vbNewLine
      SQL = SQL & " Order By DSCMODU"
      
      SQL = ""
      SQL = SQL & "Select M.*" & vbNewLine
      SQL = SQL & " From MODULO M LEFT JOIN MODULO_SISTEMA S On M.ID=S.ID"
      SQL = SQL & " Where M.VBSCRIPT like '%.NG_CADASTRO.%'" & vbNewLine
      SQL = SQL & " And S.CODSIS='" & mvarSys.CODSIS & "'" & vbNewLine
      SQL = SQL & " And M.SITMODU='S'" & vbNewLine
      SQL = SQL & " And isNull(S.MENU, M.MENUDEFAULT)='S' " & vbNewLine
      SQL = SQL & " Order By DSCMODU"
      If mvarSys.Xdb.AbreTabela(SQL, MyRs) Then
         While Not MyRs.EOF
            Call AddTaskPanelItem(Group, MyRs("IDMODU"), MyRs("DSCMODU"))
            MyRs.MoveNext
         Wend
      Else
         Call AddTaskPanelItem(Group, "CADOAREA", "Área de Aplicação")
         Call AddTaskPanelItem(Group, "CADFDESPESA", "Categoria da Despesa")
         Call AddTaskPanelItem(Group, "CADOCLASSE", "Classe de Contatos")
         Call AddTaskPanelItem(Group, "CADCCORRENTE", "Conta Corrente")
         Call AddTaskPanelItem(Group, "CADCPROMOCAO", "Descontos/Promoções")
         Call AddTaskPanelItem(Group, "CADRFUNCIONARIO", "Funcionário")
         Call AddTaskPanelItem(Group, "CADOMAQ", "Máquina")
         Call AddTaskPanelItem(Group, "CADOMANIPULO", "Manípulo")
         Call AddTaskPanelItem(Group, "CADSPRODUTO", "Produtos")
         Call AddTaskPanelItem(Group, "CADOTPCONHEC", "Tipo de Conhecimento")
         Call AddTaskPanelItem(Group, "CADOTPMAQ", "Tipo de Máquina")
         Call AddTaskPanelItem(Group, "CADOTPSERVICO", "Tipo de Serviço")
         Call AddTaskPanelItem(Group, "CADOTPTRATAMENTO", "Tipo de Tratamento")
      End If
      If .Groups.Count = 1 Then
         .Groups(1).Expandable = False
      End If
      
      .Reposition
   End With
End Sub
Private Sub AddTaskPanelItem(ByRef pGroup As TaskPanelGroup, pIDMODU As String, pCaption As String)
   Dim Item    As TaskPanelGroupItem
   Dim nId     As Integer
   
   If mvarSys.Xdb.AbreTabela("Select ID From MODULO Where IDMODU='" & pIDMODU & "'") Then
      nId = xVal(mvarSys.Xdb.RsAux("ID") & "")
   Else
      nId = pGroup.Items.Count + 1
   End If
   
   
   Set Item = pGroup.Items.Add(nId, pCaption, xtpTaskItemTypeLink, 2)
   Call SetTag(Item, "IDMODU", pIDMODU)
               
   Dim SQL As String
   Dim MyRs As Object
   SQL = "Select C.ID, M.IDMODU, C.GRUPO, M.DSCMODU, C.CONSULTA"
   SQL = SQL & ", M.VBSCRIPT, C.TL_CAD, C.TAGCAMPOS, IsNull(U.IDUSU, 'DIO') [IDUSU]"
   SQL = SQL & " From  GCAD C Join MODULO M On C.ID=M.ID"
   SQL = SQL & " Join MODULO_SISTEMA S On S.ID=M.ID"
   SQL = SQL & " Left Join GRPACESSO G On M.ID=G.ID"
   SQL = SQL & " Left Join USU_GRUPOS U On G.IDGRUPO=G.IDGRUPO"
   SQL = SQL & " Where M.SITMODU = " & SqlStr("S")
   SQL = SQL & " And S.CODSIS = " & SqlStr(mvarSys.CODSIS)
   SQL = SQL & " AND IsNull(U.IDUSU, 'DIO') =" & SqlStr(mvarSys.IDUSU)
   SQL = SQL & " AND M.IDMODU =" & SqlStr(pIDMODU)
   SQL = SQL & " Order By C.GRUPO, M.DSCMODU"
   If mvarSys.Xdb.AbreTabela(SQL, MyRs) Then
      'While Not MyRS.EOF
      Call SetTag(Item, "IDMODU", MyRs("IDMODU") & "")
      Call SetTag(Item, "CONSULTA", MyRs("CONSULTA") & "")
      Call SetTag(Item, "VBSCRIPT", MyRs("VBSCRIPT") & "")
      Call SetTag(Item, "TL_CAD", MyRs("TL_CAD") & "")
      Call SetTag(Item, "TAGCAMPOS", Encrypt2(MyRs("TAGCAMPOS") & ""))
         'MyRS.MoveNext
      'Wend
   End If
End Sub
Public Sub Show()
   '* Carregar e Acoplar Form no MDI
   Load mvarMe
   Call AcoplarForm(mvarMe, 1, mvarSys, pMDI:=mvarMDI) ' mvarSys.ObjAux)
   'mvarMe.Show
   
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
   Call ItemClick(Item)
End Sub
Public Sub ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem, Optional pReload As Boolean)
   Dim SQL     As String
   Dim sScript As String
   Dim sTl_Cad As String
   Dim sTag    As String
   
   Dim oTLCAD  As Object
   Dim sIDMODU As String
         
   Set mvarTLCadastro = Nothing
   Set mvarTLCadastro = New TL_Cadastro

   mvarItemCad = Item.Id

   sIDMODU = Trim(GetTag(Item, "IDMODU", ""))
   mvarTL_CAD = "TL_" & sIDMODU
   Set oTLCAD = CriarObjeto("CADASTRO3R." & mvarTL_CAD, False)
   'If oTLCAD Is Nothing Then
   '   Set oTLCAD = GetCadGenerico(mvarTL_CAD)
   'End If

   SQL = Trim(GetTag(Item, "CONSULTA", ""))
   sScript = Trim(GetTag(Item, "VBSCRIPT", ""))
   sTl_Cad = Trim(GetTag(Item, "TL_CAD", ""))
   sTag = Trim(Decrypt2(GetTag(Item, "TAGCAMPOS", "")))
   sIDMODU = Trim(GetTag(Item, "IDMODU", ""))
      

   With mvarTLCadastro
      Set .Sys = mvarSys
      Set .MDI = mvarMDI
   
      If Not oTLCAD Is Nothing Then
         Set oTLCAD.Sys = mvarSys
         Set oTLCAD.TLCadastro = mvarTLCadastro
      End If
            
            
      If Trim(SQL) <> "" Then .Consulta = SQL
      If Trim(Item.Caption) <> "" Then .Titulo = Item.Caption
      If Trim(sTl_Cad) <> "" Then .TL_CAD = sTl_Cad
      If Trim(sTag) <> "" Then .TAGCAMPOS = sTag
      If Trim(sIDMODU) <> "" Then .IDMODU = sIDMODU
      
      .Show pReload
   End With
End Sub
Private Function GetCadGenerico(mvarTL_CAD) As Object
   Dim MyCad As TL_CADGenerico
   Set MyCad = New TL_CADGenerico
   With MyCad
      Set .Sys = mvarSys
      'Set .RptRec = mvarTLCadastro.oTLCAD.GrdCadastro.SelectedRows(0)
      Set .ShortBar = Me
      
      .Id = "IDTPSERVICO"
      .TbName = "OTPSERVICO"
      Set .TBObj = CriarObjeto("BANCO_3R.TB_OTPSERVICO")
   End With
End Function
