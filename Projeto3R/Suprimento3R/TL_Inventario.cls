VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Inventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const StrPesquisa = "Pesquisar Produto"
Private WithEvents mvarMe     As FrmInventario
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
   Set mvarMe = New FrmInventario
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
   mvarMe.GrdProd.Height = 100
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

Private Sub mvarMe_GrdProdBeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
   If Not Item Is Nothing Then
      If InArray(Item.Index, Array(iCampo("VLULTCOMPRA"), iCampo("VLMEDIO"))) Then
         Item.Format = "R$ %s"
      ElseIf InArray(Item.Index, Array(iCampo("Saldo"))) Then
         If Item.Value <= Row.Record(iCampo("QTDMIN")).Value Then
            Item.ForeColor = vbRed
         Else
            Item.ForeColor = vbBlack
         End If
      End If
   End If
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
      .SccTit.Caption = "Inventário de Produtos"
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
   
   Sql = ""
   Sql = Sql & "Select P.IDPROD, P.CODPROD, P.NMPROD, P.UNIDCONTROLE, P.QTDMIN, P.QTDMAX, P.QTDCOMPRA" & vbNewLine
   Sql = Sql & ", Case P.EVENDA When 1 Then 'Venda' Else 'Consumo' End  [Objetivo]" & vbNewLine
   Sql = Sql & ", P.VLULTCOMPRA " & vbNewLine
   Sql = Sql & ", P.VLMEDIO " & vbNewLine
   Sql = Sql & ", IsNull(Sum(M.QTDITEM),0) [Saldo]" & vbNewLine
   Sql = Sql & "From SPRODUTO P" & vbNewLine
   Sql = Sql & "Left Join SMOVEST M On P.IDLOJA=M.IDLOJA And M.IDPROD=P.IDPROD" & vbNewLine
   Sql = Sql & " Where P.IDLOJA= " & mvarIDLOJA & vbNewLine
   Sql = Sql & "And P.ESERVICO=0" & vbNewLine
   Sql = Sql & "And P.SITPROD='A'" & vbNewLine
   Sql = Sql & "Group By P.IDPROD, P.CODPROD, P.NMPROD, P.UNIDCONTROLE, P.QTDMIN" & vbNewLine
   Sql = Sql & ", P.QTDMAX, P.QTDCOMPRA, P.EVENDA, P.VLVENDA, P.VLULTCOMPRA" & vbNewLine
   Sql = Sql & ", P.VLMEDIO" & vbNewLine
   Sql = Sql & "Order By P.EVENDA Desc , P.NMPROD, P.IDPROD"

   Screen.MousePointer = vbHourglass
   With mvarMe.GrdProd
      If ExisteItem(iCampo, "Objetivo") Then
         .GroupsOrder.Add .Columns(iCampo("Objetivo"))
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
      .GroupsOrder.Add .Columns(iCampo("Objetivo"))
      '.GroupsOrder.IndexOf .Columns(iCampo("Objetivo"))
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
         With mvarMe.GrdProd.Columns(i)
            .Visible = Not InArray(RsProd.Fields(i).Name, Array("IDPROD", "Objetivo", "QTDMAX", "VLMEDIO", "VLULTCOMPRA"))
         End With
         Call SetTag(mvarMe.GrdProd, UCase(RsProd.Fields(i).Name), i)
      End If
   Next
End Sub
Private Sub DefineGrdAlinha()
   Dim vCol As ReportColumn
   Dim i As Integer
   With mvarMe.GrdProd
      .Columns(iCampo("IDPROD")).HeaderAlignment = xtpAlignmentLeft
      .Columns(iCampo("CODPROD")).HeaderAlignment = xtpAlignmentLeft
      .Columns(iCampo("NMPROD")).HeaderAlignment = xtpAlignmentLeft
      .Columns(iCampo("UNIDCONTROLE")).HeaderAlignment = xtpAlignmentCenter
      .Columns(iCampo("QTDMIN")).HeaderAlignment = xtpAlignmentCenter
      .Columns(iCampo("QTDMAX")).HeaderAlignment = xtpAlignmentCenter
      .Columns(iCampo("QTDCOMPRA")).HeaderAlignment = xtpAlignmentCenter
      .Columns(iCampo("VLULTCOMPRA")).HeaderAlignment = xtpAlignmentRight
      .Columns(iCampo("VLMEDIO")).HeaderAlignment = xtpAlignmentRight
      
      .Columns(iCampo("Objetivo")).HeaderAlignment = xtpAlignmentCenter
      .Columns(iCampo("Saldo")).HeaderAlignment = xtpAlignmentRight
          
      For Each vCol In .Columns
         vCol.Alignment = vCol.HeaderAlignment
         vCol.Sortable = True
      Next
      
      .Columns(iCampo("CODPROD")).Caption = "Código"
      .Columns(iCampo("NMPROD")).Caption = "Nome"
      .Columns(iCampo("UNIDCONTROLE")).Caption = "Unid."
      .Columns(iCampo("QTDMIN")).Caption = "Mínimo."
      .Columns(iCampo("QTDMAX")).Caption = "Máximo"
      .Columns(iCampo("QTDCOMPRA")).Caption = "Compra"
      .Columns(iCampo("VLULTCOMPRA")).Caption = "Valor Compra"
      .Columns(iCampo("VLMEDIO")).Caption = "Valor Médio"
                  
      .Columns(iCampo("CODPROD")).Width = 60
      .Columns(iCampo("NMPROD")).Width = 120
      .Columns(iCampo("UNIDCONTROLE")).Width = 60
      .Columns(iCampo("QTDMIN")).Width = 80
      .Columns(iCampo("QTDMAX")).Width = 80
      .Columns(iCampo("QTDCOMPRA")).Width = 80
      .Columns(iCampo("VLULTCOMPRA")).Width = 80
      .Columns(iCampo("VLMEDIO")).Width = 80
            
      .Columns(iCampo("Objetivo")).Width = 60
      .Columns(iCampo("Saldo")).Width = 80
      .Columns(iCampo("NMPROD")).Width = (.Width / Screen.TwipsPerPixelX)
      For i = 0 To .Columns.Count - 1
         If iCampo("NMPROD") <> i And .Columns(i).Visible Then
          .Columns(iCampo("NMPROD")).Width = .Columns(iCampo("NMPROD")).Width - .Columns(i).Width
         End If
      Next
      .Columns(iCampo("NMPROD")).Width = .Columns(iCampo("NMPROD")).Width - 30
      .AutoColumnSizing = False
      '.Populate
   End With
End Sub
Private Sub mvarMe_Rezise()
   mvarMe.SccTit.Move 0, 50, mvarMe.ScaleWidth - 50, 420
   With mvarMe.txtFiltrar
      .Left = mvarMe.Width - 2800
      .Top = 90
      .Visible = False
   End With
   With mvarMe.imgLupa
      .Left = mvarMe.Width - 680
      .Top = 100
      .Height = mvarMe.txtFiltrar.Height - 25
      .Visible = False
   End With
   With mvarMe.GrdProd
      .Left = 0
      .Top = mvarMe.SccTit.Top + mvarMe.SccTit.Height + 30
      .Width = mvarMe.ScaleWidth - 50
      .Height = mvarMe.ScaleHeight - .Top - 60
   End With
End Sub

Private Sub mvarPane_WndTaskPanelItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
   If Item.Id <= 3 Then Exit Sub
   Select Case Item.Id
      Case 7: Call ExibirAviso("Entrada de estoque não disponível")
      Case 8: Call ExibirAviso("Saída de estoque não disponível")
      Case 9: Call ExibirAviso("Contagem de estoque não disponível")
   End Select
End Sub
