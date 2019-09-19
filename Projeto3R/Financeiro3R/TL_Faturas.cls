VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Faturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe     As FrmFaturas
Attribute mvarMe.VB_VarHelpID = -1

Private Enum eIconId
   Novo = 1
   Editar
   RPS
   NF
   ExpRPS
   ImpNF
   Sair
End Enum

Dim iCampo As Collection

Private RsCad           As Object
Private mvarSys         As Object
Private mvarConsulta    As String
Private mvarIDLOJA      As Integer
Public Sub F_Refresh()
   Call PopulaGrdITEM(True)
End Sub
Public Property Let IDLOJA(ByVal vData As Integer)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
End Property
Public Property Get Sys() As Object
   Set Sys = mvarSys
End Property
Public Property Get Consulta() As String
   Consulta = mvarConsulta
End Property
Public Property Let Consulta(ByVal vData As String)
   mvarConsulta = vData
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmFaturas
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_CommandBarsExecute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   Select Case Control.Id
      Case eIconId.Novo:   Call NovoItem
      Case eIconId.Editar: Call EditarItem(mvarMe.GrdITEM.SelectedRows(0).Record(iCampo("Nº FATURA")).Value)
      'Case eIconId.RPS:    Call EditarRPS
      'Case eIconId.NF:     Call EditarNF
      'Case eIconId.ExpRPS: Call ExportarRPS
      'Case eIconId.ImpNF:  Call ImportarNF
      Case eIconId.Sair:   Unload mvarMe
      Case "Sair":
   End Select
End Sub
Private Sub mvarMe_GrdITEMBeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
   If Not iCampo Is Nothing Then
      'If Item.Index = iCampo("Depósito") Or Item.Index = iCampo("Pagamento") Or Item.Index = iCampo("Saldo") Then
      '   Item.Value = IIf(xVal(Item.Value & "") = 0, "", ValBr(Item.Value & ""))
      'End If
   End If
   If ((Row.Index Mod 2 = 0) And (Row.GroupRow = False)) Then
       Metrics.BackColor = RGB(245, 245, 245)
   Else
      'Metrics.BackColor = RGB(65, 165, 241)
   End If

End Sub

Private Sub mvarMe_GrdITEMKeyUp(KeyCode As Integer, Shift As Integer)
   Call mvarMe_KeyUp(KeyCode, Shift)
End Sub

Private Sub mvarMe_GrdITEMRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   Call EditarItem(mvarMe.GrdITEM.SelectedRows(0).Record(iCampo("Nº FATURA")).Value)
End Sub
Private Sub mvarMe_Activate()
   mvarMe_Resize
   If GetTag(mvarMe, "1Vez", 0) = 0 Then
      Call SetTag(mvarMe, "1Vez", 1)
      Call PosicionaForm
   End If

   'Call MontarToolbar

   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF5 Then
      Call PopulaGrdITEM
   ElseIf KeyCode = vbKeyEscape Then
         mvarMe.txtFiltrar.Text = mvarMe.txtFiltrar.Tag
         Call PopulaGrdITEM
   ElseIf (Asc(UCase(Chr(KeyCode))) >= Asc("A") And Asc(UCase(Chr(KeyCode))) >= Asc("Z")) Or (Asc(UCase(Chr(KeyCode))) >= Asc("0") And Asc(UCase(Chr(KeyCode))) >= Asc("9")) Then
      If Not mvarMe.ActiveControl Is mvarMe.txtFiltrar Then
         mvarMe.txtFiltrar.Text = UCase(Chr(KeyCode))
         mvarMe.txtFiltrar.SetFocus
         mvarMe.txtFiltrar.SelStart = Len(mvarMe.txtFiltrar.Text)
      End If
   End If
End Sub
Private Sub mvarMe_Load()
   mvarMe.txtFiltrar.Tag = mvarMe.txtFiltrar.Text
   
   Set mvarSys.ObjAux = mvarMe
   Call PosicionaForm
   
   Call MontarTela
   'Call MontarMenu
   Call MontarToolbar
   'Call MontarStatusBar
   Call PopulaTela
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub MontarTela()
   'ReportControlGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   With mvarMe
'      With .SccContato
'         .VisualTheme = xtpShortcutThemeOffice2003
'         .GradientHorizontal = False
'      End With
      With .GrdITEM
         .BorderStyle = xtpBorderFrame
         .PaintManager.ColumnStyle = xtpColumnOffice2007
         .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
         .PaintManager.RefreshMetrics
         .Redraw
      End With
   End With
End Sub
Private Sub PopulaTela()
   Call PopulaGrdITEM
   mvarMe_Resize
End Sub
Private Sub mvarMe_Resize()
   On Error Resume Next
   'mvarMe.SccContato.Move 0, 0, mvarMe.ScaleWidth, 400

   mvarMe.GrdITEM.Left = 0
   'mvarMe.GrdITEM.Top = mvarMe.SccContato.Top + mvarMe.SccContato.Height
   mvarMe.GrdITEM.Top = 400
   mvarMe.GrdITEM.Width = mvarMe.ScaleWidth - 50
   mvarMe.GrdITEM.Height = mvarMe.ScaleHeight - mvarMe.GrdITEM.Top - 50

   mvarMe.txtFiltrar.Left = mvarMe.Width - 2800
   mvarMe.imgLupa.Left = mvarMe.Width - 680

   'mvarMe.txtFiltrar.Top = (mvarMe.SccContato.Height - mvarMe.txtFiltrar.Height) / 2
   mvarMe.txtFiltrar.Top = (400 - mvarMe.txtFiltrar.Height) / 2
   mvarMe.imgLupa.Top = mvarMe.txtFiltrar.Top + 5
   mvarMe.imgLupa.Height = mvarMe.txtFiltrar.Height - 30
   
   mvarMe.txtFiltrar.ZOrder 0
   mvarMe.imgLupa.ZOrder 0
End Sub
Private Sub PosicionaForm()
   Dim Left As Long
   Dim Top As Long
   Dim Right As Long
   Dim Bottom As Long
   On Error Resume Next
   mvarSys.MDI.CommandBars.GetClientRect Left, Top, Right, Bottom
   mvarSys.MDI.CommandBars(1).GetWindowRect Left, Top, Right, Bottom
      
   mvarMe.Top = Bottom
   
   mvarMe.Left = mvarSys.MDI.Left + 60
   mvarMe.Width = mvarSys.MDI.Width - mvarMe.Left - 120
   mvarMe.Height = mvarSys.MDI.Height - mvarMe.Top - 420
End Sub
Private Sub mvarMe_txtFiltrarGotFocus()
   mvarMe.txtFiltrar.ForeColor = vbBlack
   If mvarMe.txtFiltrar.Text = mvarMe.txtFiltrar.Tag Then
      mvarMe.txtFiltrar.Text = ""
   End If
End Sub
Private Sub mvarMe_txtFiltrarKeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      mvarMe.GrdITEM.FilterText = mvarMe.txtFiltrar.Text
      mvarMe.GrdITEM.Populate
   End If
End Sub
Private Sub mvarMe_txtFiltrarLostFocus()
   If mvarMe.txtFiltrar.Text = "" Then mvarMe.txtFiltrar.Text = mvarMe.txtFiltrar.Tag
   mvarMe.txtFiltrar.ForeColor = 12632256
End Sub
Public Sub PopulaGrdITEM(Optional bRefresh As Boolean)
   Dim i As Long
   Dim nRow As Long
   Dim Sql  As String
      
   With mvarMe.GrdITEM
      .SetImageList mvarMe.ImageList1
     
      If bRefresh Then
         nRow = .SelectedRows(0).Index
      End If
      If mvarConsulta = "" Then
         Sql = QrySelect
      Else
         Sql = mvarConsulta
      End If
            
      If mvarSys.xDb.AbreTabela(Sql, RsCad) Then
         Call FillRCFromRS(RsCad, mvarMe.GrdITEM, pCollColumn:=iCampo, pCurrency:=True)
      Else
         mvarMe.CommandBars.ActiveMenuBar.Controls(eIconId.Editar).Enabled = False
      End If
      
      For i = 0 To .Records.Count - 1
         If .Records(i).Item(0).Value = 0 Then
            .Records(i).Item(0).Icon = 0
         ElseIf .Records(i).Item(0).Value = 10 Then
            .Records(i).Item(0).Icon = 1
         End If
         .Records(i).Item(0).BackColor = &HE0E0E0
         .Records(i).Item(0).ForeColor = &HE0E0E0
         '.Populate
      Next
      If .Columns.Count > 0 Then
         .Columns(0).Width = 18
         .Columns(0).Caption = ""
         .Columns(0).Icon = 1
      End If
      If mvarMe.txtFiltrar.Text = "" Or mvarMe.txtFiltrar.Text = mvarMe.txtFiltrar.Tag Then
         .FilterText = ""
      Else
         .FilterText = mvarMe.txtFiltrar.Text
      End If
      
      
      .Populate
      If bRefresh Then
         If nRow > .Rows.Count - 1 Then
            nRow = .Rows.Count - 1
         End If
         .Rows(nRow).EnsureVisible
         .Rows(nRow).Selected = True
      End If

'      Call DefineGrdIcones
'      Call DefineGrdFiltro
'      Call CalculaSaldo
      Call DefineGrdAlinha
      .Redraw
   End With

End Sub
Private Function QrySelect()
   Dim Sql As String
   
   Sql = "Select SITFATURA"
   Sql = Sql & ", Right('000000'+Rtrim(Cast(F.IDFATURA As Char)),6) [Nº Fatura]"
   Sql = Sql & ", Cast(Convert(Char(10),F.DTEMISSAO, 103) As SmallDatetime) [Emissão]"
   Sql = Sql & ", F.VALOR [Valor]"
   Sql = Sql & ", Cast(Convert(Char(10),F.DTPREV, 103) As SmallDatetime) [Vencimento]"
   Sql = Sql & ", Cast(Convert(Char(10),V.DTVENDA, 103) As SmallDatetime) [Baixa]"
   Sql = Sql & ", Right('000000'+Rtrim(Cast(V.IDVENDA As Char)),6) [Nº Venda]"
   Sql = Sql & ", Right('000000'+Rtrim(Cast(N.NUMNOTA As Char)),6) [Nº N.F.]"
   Sql = Sql & ", C.NOME [Cliente]"
   Sql = Sql & " From FFATURA F"
   Sql = Sql & " Join OCLIENTE C          On F.IDLOJA=C.IDLOJA And F.IDCLIENTE=C.IDCLIENTE"
   Sql = Sql & " Left Join CVENDA V       On F.IDLOJA=V.IDLOJA And F.IDVENDA=V.IDVENDA"
   Sql = Sql & " Left Join FNOTAFISCAL N  On V.IDLOJA=N.IDLOJA And V.IDVENDA=N.IDVENDA And N.FLGCANCELADA=0"
   Sql = Sql & " Where F.IDLOJA=" & SqlNum(mvarIDLOJA)
   'Sql = Sql & " And V.SITVENDA='10'"
   'Sql = Sql & " And V.VLVENDA-V.VLDESC>0"
   Sql = Sql & " Order By Emissão desc, F.IDFATURA desc"
   
   QrySelect = Sql
End Function
Private Sub DefineGrdAlinha()
   Dim vCol As ReportColumn
   If iCampo Is Nothing Then Exit Sub
   With mvarMe.GrdITEM
      For Each vCol In .Columns
         vCol.HeaderAlignment = xtpAlignmentCenter
         'vCol.Sortable = False
      Next
      .Columns(iCampo("Cliente")).HeaderAlignment = xtpAlignmentLeft
      If iCampo.Count <> 0 Then
         .Columns(iCampo("Nº Fatura")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Emissão")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Valor")).Alignment = xtpAlignmentRight
         .Columns(iCampo("Vencimento")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Baixa")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Nº Venda")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Nº N.F.")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Cliente")).Alignment = xtpAlignmentLeft
         
         .AutoColumnSizing = False
         .Columns(iCampo("Nº Fatura")).Width = 60
         .Columns(iCampo("Emissão")).Width = 90
         .Columns(iCampo("Valor")).Width = 100
         .Columns(iCampo("Vencimento")).Width = 90
         .Columns(iCampo("Baixa")).Width = 90
         .Columns(iCampo("Nº Venda")).Width = 60
         .Columns(iCampo("Nº N.F.")).Width = 60
         .Columns(iCampo("Cliente")).Width = 400 '(.Width / Screen.TwipsPerPixelX) - 420
         '.Populate
      End If
   End With
End Sub

Private Sub MontarToolbar()
   Dim MyToolBar As CommandBar
   
   With mvarMe.CommandBars
'      .DeleteAll
      .EnableCustomization False
      Set .Icons = mvarMe.ImgToobar.Icons
      Set MyToolBar = mvarMe.CommandBars.ActiveMenuBar
      With MyToolBar
         .EnableDocking xtpFlagStretched
         '.DefaultButtonStyle = xtpButtonAutomatic '= xtpButtonIcon
         '.Customizable = True
         '.Closeable = False
         '.Position = xtpBarTop
         '.ShowExpandButton = False
         '.ShowGripper = True
         '.ShowTextBelowIcons = False
         '.ContextMenuPresent = True
            
         Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.Novo, pIconId:=10001, pCaption:="Novo")
         Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.Editar, pIconId:=10002, pCaption:="Editar")
         'Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.RPS, pIconId:=10003, pCaption:="RPS", pBeginGroup:=True)
         'Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.NF, pIconId:=10004, pCaption:="N.F.")
         'Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.ExpRPS, pIconId:=10005, pCaption:="Exportar RPS", pBeginGroup:=True)
         'Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.ImpNF, pIconId:=10006, pCaption:="Importar NF")
         Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.Sair, pIconId:=10007, pCaption:="Sair", pBeginGroup:=True)
         
         Dim n As Object
         For Each n In .Controls
            n.Category = "BAR_VENDA"
         Next
         
         .Position = 0
         .Visible = True
      End With
   End With
End Sub
Private Sub NovoItem()
   Dim MyItem As TL_Fatura
   
   Set MyItem = New TL_Fatura
   With MyItem
      Set .Sys = mvarSys
      .Show vbModal
   End With
   Set MyItem = Nothing
End Sub
Private Sub EditarItem(pIDITEM As Long)
   Dim MyItem As TL_Fatura
   Dim nRow As Long
   
   Screen.MousePointer = vbHourglass
   
   Set MyItem = New TL_Fatura
   With MyItem
      Set .Sys = mvarSys
      .IDLOJA = mvarIDLOJA
      .IDFATURA = pIDITEM
      Set .TlFaturas = Me
      .Show vbModal
      
      If .IDVENDA <> 0 Then
         If mvarMe.GrdITEM.SelectedRows.Count > 0 Then
            nRow = mvarMe.GrdITEM.SelectedRows(0).Index
         End If
         Call PopulaGrdITEM
         
         If mvarMe.GrdITEM.SelectedRows.Count > 0 Then
            mvarMe.GrdITEM.Rows(mvarMe.GrdITEM.SelectedRows(0).Index).Selected = False
         End If
         mvarMe.GrdITEM.Rows(nRow).Selected = True
         mvarMe.GrdITEM.Redraw
      End If
   End With
   Screen.MousePointer = vbDefault
   Exit Sub
Saida:
   Screen.MousePointer = vbDefault
   ShowError
End Sub
