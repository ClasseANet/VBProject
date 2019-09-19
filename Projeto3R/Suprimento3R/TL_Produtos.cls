VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Produtos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const StrPesquisa = "Pesquisar Produto"
Private WithEvents mvarMe     As FrmProdutos
Attribute mvarMe.VB_VarHelpID = -1
Private WithEvents mvarPane   As FrmPaneSupri
Attribute mvarPane.VB_VarHelpID = -1

Private RsProd As Recordset  'Object
Dim iCampo As Collection

Private mvarSys      As Object
Private mvarIDLOJA   As Integer
Public Property Set Pane(ByVal vData As Object)
   Set mvarPane = vData
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
End Property
Public Property Get Sys() As Object
   Set Sys = mvarSys
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Sub Show(Optional Reload As Boolean = False)
   '* Carregar e Acoplar Form no MDI
   Load mvarMe
   Call AcoplarForm(mvarMe, 2, mvarSys)
   mvarMe.Show
   'mvarMe_Resize
   
   '* Definir foco
   With mvarMe.GrdProd
      If .Visible And .Enabled Then
         .SetFocus
      End If
   End With
End Sub
Public Sub Hide()
   mvarMe.Hide
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmProdutos
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_Activate()
   'Call MontaTela
   'Call PopulaTela
   If GetTag(mvarMe, "LOAD", 1) = 1 Then
      Call SetTag(mvarMe, "LOAD", 0)
      Call DefineGrdAlinha
   End If
   mvarMe.GrdProd.Width = mvarMe.ScaleWidth + 100
   mvarMe.GrdProd.Height = mvarMe.ScaleHeight + 100
   mvarMe_Rezise
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarMe_CmdEditarClick()
   Dim mvarTLOnFocus As Object
   Dim nIndex  As Long
   Dim nTopRow As Long
   
   If mvarTLOnFocus Is Nothing Then
      Set mvarTLOnFocus = CriarObjeto("CADASTRO3R.TL_CADSPRODUTO", False)
   End If
   Set mvarTLOnFocus.Sys = mvarSys
   'Set mvarTLOnFocus.ShortBar = mvarPane
   Set mvarTLOnFocus.RptRec = mvarMe.GrdProd.SelectedRows(0)
   
   nTopRow = mvarMe.GrdProd.TopRowIndex
   nIndex = mvarMe.GrdProd.SelectedRows(0).Index
   mvarTLOnFocus.Show
   Call PopulaTela
   On Error Resume Next
   mvarMe.GrdProd.TopRowIndex = nTopRow
   mvarMe.GrdProd.Rows(nIndex).EnsureVisible
   mvarMe.GrdProd.Rows(nIndex).Selected = True
   
   Set mvarTLOnFocus = Nothing
End Sub

Private Sub mvarMe_CmdExcluirClick()
   Dim mvarTLOnFocus As Object
   Dim nIndex  As Long
   Dim nTopRow As Long
   
   If mvarTLOnFocus Is Nothing Then
      Set mvarTLOnFocus = CriarObjeto("CADASTRO3R.TL_CADSPRODUTO", False)
   End If
   Set mvarTLOnFocus.Sys = mvarSys
   'Set mvarTLOnFocus.ShortBar = mvarPane
   Set mvarTLOnFocus.RptRec = mvarMe.GrdProd.SelectedRows(0)
   Call mvarTLOnFocus.Excluir
   Call PopulaTela
   On Error Resume Next
   mvarMe.GrdProd.TopRowIndex = nTopRow
   mvarMe.GrdProd.Rows(nIndex).EnsureVisible
   mvarMe.GrdProd.Rows(nIndex).Selected = True
   
   
   
   Set mvarTLOnFocus = Nothing
End Sub

Private Sub mvarMe_CmdNovoClick()
   Dim mvarTLOnFocus As Object
   Dim nIndex As Long
   Dim nTopRow  As Long
   
   If mvarTLOnFocus Is Nothing Then
      Set mvarTLOnFocus = CriarObjeto("CADASTRO3R.TL_CADSPRODUTO", False)
   End If
   Set mvarTLOnFocus.Sys = mvarSys
   
   nTopRow = mvarMe.GrdProd.TopRowIndex
   nIndex = mvarMe.GrdProd.SelectedRows(0).Index
   mvarTLOnFocus.Show
   Call PopulaTela
   On Error Resume Next
   mvarMe.GrdProd.TopRowIndex = nTopRow
   mvarMe.GrdProd.Rows(nIndex).EnsureVisible
   mvarMe.GrdProd.Rows(nIndex).Selected = True
   
   Set mvarTLOnFocus = Nothing
End Sub

Private Sub mvarMe_GrdProdKeyUp(KeyCode As Variant, Shift As Variant)
   If KeyCode = vbKeyF5 Then Call PopulaTela
End Sub
Private Sub mvarMe_GrdProdRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   Call mvarMe_CmdEditarClick
End Sub
Private Sub mvarMe_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyF5 Then Call PopulaTela
End Sub
Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF5 Then Call PopulaTela
End Sub
Private Sub mvarMe_Load()
   Call MontaTela
   Call PopulaTela
End Sub
Private Sub MontaTela()
   'ReportControlGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   
   With mvarMe
      .txtFiltrar = StrPesquisa
      .SccTit.Caption = "Lista de Produtos/Serviços"
      .SccTit.VisualTheme = xtpShortcutThemeOffice2003
      .SccTit.GradientHorizontal = False
      
      With .GrdProd
         .BorderStyle = xtpBorderFrame
         .PaintManager.ColumnStyle = xtpColumnOffice2007
         .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
         .PaintManager.RefreshMetrics
         
         '.SetImageList mvarMe.ImageList1
         .SetCustomDraw xtpCustomBeforeDrawRow
         
         .MultipleSelection = False
         .AllowColumnRemove = False
         .AllowColumnReorder = False
         .AllowColumnResize = True
         .AllowColumnSort = True
         .AllowEdit = False
                     
         .FooterRowsAllowAccess = True
         .FooterRowsAllowEdit = True
         .FooterRowsEnableSelection = True
      End With
   End With
   Call mvarMe_Rezise
End Sub
Private Sub PopulaTela()
   Dim Sql As String
   Dim i As Long
   Dim nTopRow As Long
   Dim oRecord As ReportRecord
   Dim oRptRow As ReportRow
   Sql = "Select * "
   Sql = Sql & ", Case ESERVICO When 1 Then 'Serviço' Else 'Material' End  [Tipo]"
   Sql = Sql & ", Case EVENDA When 1 Then 'Venda' Else 'Consumo' End  [Objetivo]"
   Sql = Sql & ", 'R$ '+Replace(Cast(VLVENDA as varchar), '.', ',') [Valor]"
   'Sql = Sql & ", VLULTCOMPRA [Valor]"
   'Sql = Sql & ", VLMEDIO [Valor Médio]"
   Sql = Sql & " From SPRODUTO"
   Sql = Sql & " Where IDLOJA= " & mvarIDLOJA
   Sql = Sql & " Order By ESERVICO Desc, EVENDA Desc , NMPROD, IDPROD"
   
   
   Screen.MousePointer = vbHourglass
   With mvarMe.GrdProd
      If ExisteItem(iCampo, "Tipo") Then
         .GroupsOrder.Add .Columns(iCampo("Tipo"))
      End If
'      .Visible = False
      Call mvarSys.xDb.AbreTabela(Sql, RsProd)
      Call FillRCFromRS(RsProd, mvarMe.GrdProd)
      Call EscondeCampos
      'Call DefineGrdFiltro
      Call DefineGrdAlinha
      'Call InserirGrdLinha
      
      .GroupsOrder.DeleteAll
      .ShowGroupBox = False
      .ShowItemsInGroups = False
      .SkipGroupsFocus = False
      .GroupsOrder.Add .Columns(iCampo("Tipo"))
      '.GroupsOrder.IndexOf .Columns(iCampo("Tipo"))
      .Visible = True
      
      .Populate
      .Redraw
      If .Visible And .Enabled Then
         .SetFocus
      End If
      If .Rows.Count > 0 Then
         .Rows(0).EnsureVisible
         .Rows(0).Selected = True
         .Redraw
      
         '.Rows(0).Selected = True
'         .Rows(.TopRowIndex).EnsureVisible
'         .Populate
'         .GroupsOrder(0).EnsureVisible
'         .GroupsOrder(0).BestFit
      End If
      
   End With
   Screen.MousePointer = vbDefault
End Sub
Private Sub EscondeCampos()
   Dim i As Integer
   
   Set iCampo = New Collection
   
   If RsProd.State = 0 Then Exit Sub
   
   
   Call SetTag(mvarMe.GrdProd, "COUNT", RsProd.Fields.Count)
   For i = 0 To RsProd.Fields.Count - 1
      If ExisteItem(iCampo, RsProd.Fields(i).Name) Then
         mvarMe.GrdProd.Columns(i).Visible = False
      Else
         iCampo.Add i, RsProd.Fields(i).Name
         'mvarMe.GrdProd.Columns(i).Visible = InArray(RsProd.Fields(i).Name, Array("CODPROD", "NMPROD", "DSCPROD", "UNIDCONTROLE", "VLVENDA", "EVENDA", "ESERVICO", "SITPROD"))
         mvarMe.GrdProd.Columns(i).Visible = InArray(RsProd.Fields(i).Name, Array("CODPROD", "NMPROD", "DSCPROD", "UNIDCONTROLE", "SITPROD"))
         mvarMe.GrdProd.Columns(i).Visible = InArray(RsProd.Fields(i).Name, Array("Objetivo", "Valor")) Or mvarMe.GrdProd.Columns(i).Visible
         Call SetTag(mvarMe.GrdProd, UCase(RsProd.Fields(i).Name), i)
      End If
   Next
End Sub
Private Sub DefineGrdAlinha()
   Dim vCol As ReportColumn
   Dim i As Integer
   With mvarMe.GrdProd
      .Columns(iCampo("CODPROD")).HeaderAlignment = xtpAlignmentLeft
      .Columns(iCampo("NMPROD")).HeaderAlignment = xtpAlignmentLeft
      .Columns(iCampo("DSCPROD")).HeaderAlignment = xtpAlignmentLeft
      .Columns(iCampo("UNIDCONTROLE")).HeaderAlignment = xtpAlignmentCenter
      .Columns(iCampo("VLVENDA")).HeaderAlignment = xtpAlignmentRight
      .Columns(iCampo("EVENDA")).HeaderAlignment = xtpAlignmentCenter
      .Columns(iCampo("ESERVICO")).HeaderAlignment = xtpAlignmentCenter
      
      .Columns(iCampo("Tipo")).HeaderAlignment = xtpAlignmentCenter
      .Columns(iCampo("Objetivo")).HeaderAlignment = xtpAlignmentCenter
      .Columns(iCampo("Valor")).HeaderAlignment = xtpAlignmentRight
           
      .Columns(iCampo("SITPROD")).HeaderAlignment = xtpAlignmentCenter
      For Each vCol In .Columns
         vCol.Alignment = vCol.HeaderAlignment
         vCol.Sortable = True
      Next
      
      .Columns(iCampo("CODPROD")).Caption = "Código"
      .Columns(iCampo("NMPROD")).Caption = "Nome"
      .Columns(iCampo("DSCPROD")).Caption = "Descrição"
      .Columns(iCampo("UNIDCONTROLE")).Caption = "Unid."
      '.Columns(iCampo("VLVENDA")).Caption = "Valor Venda"
      '.Columns(iCampo("EVENDA")).Caption = "Venda"
      '.Columns(iCampo("ESERVICO")).Caption = "Serviço"
      .Columns(iCampo("SITPROD")).Caption = "Situação"
                  
      .Columns(iCampo("CODPROD")).Width = 60
      .Columns(iCampo("NMPROD")).Width = 120
      .Columns(iCampo("DSCPROD")).Width = 120
      .Columns(iCampo("UNIDCONTROLE")).Width = 60
      .Columns(iCampo("VLVENDA")).Width = 80
      .Columns(iCampo("EVENDA")).Width = 60
      .Columns(iCampo("ESERVICO")).Width = 60
      .Columns(iCampo("SITPROD")).Width = 60
      
      .Columns(iCampo("Tipo")).Width = 60
      .Columns(iCampo("Objetivo")).Width = 60
      .Columns(iCampo("Valor")).Width = 80
            
      .Columns(iCampo("DSCPROD")).Width = (.Width / Screen.TwipsPerPixelX)
      For i = 0 To .Columns.Count - 1
         If iCampo("DSCPROD") <> i And .Columns(i).Visible Then
          .Columns(iCampo("DSCPROD")).Width = .Columns(iCampo("DSCPROD")).Width - .Columns(i).Width
         End If
      Next
      .Columns(iCampo("DSCPROD")).Width = .Columns(iCampo("DSCPROD")).Width - 25
      
      .AutoColumnSizing = False
      '.Populate
   End With
End Sub
Private Sub mvarMe_Rezise()
   mvarMe.SccTit.Move 0, 50, mvarMe.ScaleWidth - 50, 420
   With mvarMe.txtFiltrar
      .left = mvarMe.Width - 2800
      .top = 90
      .Visible = False
   End With
   With mvarMe.imgLupa
      .left = mvarMe.Width - 680
      .top = 100
      .Height = mvarMe.txtFiltrar.Height - 25
      .Visible = False
   End With
   With mvarMe.GrdProd
      .left = 0
      .top = mvarMe.SccTit.top + mvarMe.SccTit.Height + 30
      .Width = mvarMe.ScaleWidth - 50
      .Height = mvarMe.ScaleHeight - .top - 60
   End With
End Sub

Private Sub mvarPane_WndTaskPanelItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
   If Item.Id <= 3 Then Exit Sub
   Select Case Item.Id
      Case 4: mvarMe_CmdNovoClick
      Case 5: mvarMe_CmdEditarClick
      Case 6: mvarMe_CmdExcluirClick
   End Select
End Sub
