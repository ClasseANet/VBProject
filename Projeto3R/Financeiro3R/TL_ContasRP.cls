VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_ContasRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe     As FrmContasRP
Attribute mvarMe.VB_VarHelpID = -1
Private Enum eIconId
   Novo = 1
   Editar
   Sair
End Enum

Dim iCampo As Collection

Private RsCad           As Object
Private mvarSys         As Object
Private mvarConsulta    As String
Private mvarIDLOJA      As Integer
Public Sub F_Refresh()
   Call PopulaGrdContasRP(True)
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
   Set mvarMe = New FrmContasRP
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_CommandBarsExecute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   Select Case Control.Id
      Case eIconId.Novo:   Call NovaFatura
      Case eIconId.Editar: Call EditarFatura(mvarMe.GrdContasRP.SelectedRows(0).Record(iCampo("Nº VENDA")).Value)
      Case eIconId.Sair:   Unload mvarMe
      Case "Sair":
   End Select
End Sub

Private Sub mvarMe_GrdContasRPBeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
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
Private Sub mvarMe_GrdContasRPKeyUp(KeyCode As Integer, Shift As Integer)
   Call mvarMe_KeyUp(KeyCode, Shift)
End Sub
Private Sub mvarMe_GrdContasRPRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   Call EditarFatura(mvarMe.GrdContasRP.SelectedRows(0).Record(iCampo("Nº VENDA")).Value)
End Sub
Private Sub mvarMe_Activate()
   mvarMe_Resize
   If GetTag(mvarMe, "1Vez", 0) = 0 Then
      Call SetTag(mvarMe, "1Vez", 1)
      Call PosicionaForm
   End If

   Call MontarToolbar

   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF5 Then
      Call PopulaGrdContasRP
   ElseIf KeyCode = vbKeyEscape Then
         mvarMe.txtFiltrar.Text = mvarMe.txtFiltrar.Tag
         Call PopulaGrdContasRP
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
'   Call MontarToolbar
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
      With .GrdContasRP
         .BorderStyle = xtpBorderFrame
         .PaintManager.ColumnStyle = xtpColumnOffice2007
         .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
         .PaintManager.RefreshMetrics
         .Redraw
      End With
   End With
End Sub
Private Sub PopulaTela()
   Call PopulaGrdContasRP
   mvarMe_Resize
End Sub
Private Sub mvarMe_Resize()
   On Error Resume Next
   'mvarMe.SccContato.Move 0, 0, mvarMe.ScaleWidth, 400

   mvarMe.GrdContasRP.Left = 0
   'mvarMe.GrdContasRP.Top = mvarMe.SccContato.Top + mvarMe.SccContato.Height
   mvarMe.GrdContasRP.Top = 400
   mvarMe.GrdContasRP.Width = mvarMe.ScaleWidth - 50
   mvarMe.GrdContasRP.Height = mvarMe.ScaleHeight - mvarMe.GrdContasRP.Top - 50

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
      mvarMe.GrdContasRP.FilterText = mvarMe.txtFiltrar.Text
      mvarMe.GrdContasRP.Populate
   End If
End Sub

Private Sub mvarMe_txtFiltrarLostFocus()
   
   If mvarMe.txtFiltrar.Text = "" Then mvarMe.txtFiltrar.Text = mvarMe.txtFiltrar.Tag
   mvarMe.txtFiltrar.ForeColor = 12632256
End Sub
Private Sub PopulaGrdContasRP(Optional bRefresh As Boolean)
   Dim i As Long
   Dim nRow As Long
   Dim Sql  As String
      
   With mvarMe.GrdContasRP
      .SetImageList mvarMe.ImageList1
     
      If bRefresh Then
         nRow = .SelectedRows(0).Index
      End If
      If mvarConsulta = "" Then
         Sql = "Select Right('000000'+Rtrim(Cast(C.IDCONTA As Char)),6) [Nº Conta]"
         Sql = Sql & ", Cast(Convert(Char(10), C.DTPREV, 103) As SmallDatetime) [Data]"
         Sql = Sql & ", I.NOME [Cliente]"
         Sql = Sql & ", Right('000000'+Rtrim(Cast(A.IDATENDIMENTO As Char)),6) [Atendimento]"
         Sql = Sql & ", Right('000000'+Rtrim(Cast(C.IDVENDA As Char)),6) [Venda]"
         Sql = Sql & ", C.VALOR [Valor]"
         Sql = Sql & " From FCONTAR C"
         Sql = Sql & " Join OCLIENTE I On C.IDCLIENTE=I.IDCLIENTE"
         Sql = Sql & " Left Join OATENDIMENTO A On I.IDLOJA=A.IDLOJA And I.IDATENDIMENTO=A.IDATENDIMENTO"
         Sql = Sql & " Left Join CVENDA V On I.IDLOJA=V.IDLOJA And I.IDVENDA=V.IDVENDA"
         Sql = Sql & " Where C.IDLOJA=" & SqlNum(mvarIDLOJA)
         'Sql = Sql & " And V.SITVENDA='10'"
         'Sql = Sql & " And V.VLVENDA-V.VLDESC>0"
         Sql = Sql & " Order By Data desc, C.IDCONTA desc"
      Else
         Sql = mvarConsulta
      End If
            
      If mvarSys.xDb.AbreTabela(Sql, RsCad) Then
         Call FillRCFromRS(RsCad, mvarMe.GrdContasRP, pCollColumn:=iCampo, pCurrency:=True)
      End If
      
      'For i = 0 To .Records.Count - 1
      '   .Records(i).Item(1).Icon = 0
      'Next
      
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
Private Sub DefineGrdAlinha()
   Dim vCol As ReportColumn
   If iCampo Is Nothing Then Exit Sub
   
   With mvarMe.GrdContasRP
      For Each vCol In .Columns
         vCol.HeaderAlignment = xtpAlignmentCenter
         'vCol.Sortable = False
      Next
      If iCampo.Count <> 0 Then
         .Columns(iCampo("Nº Conta")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Data")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Cliente")).Alignment = xtpAlignmentLeft
         .Columns(iCampo("Atendimento")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Venda")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Valor")).Alignment = xtpAlignmentRight
         
         .AutoColumnSizing = False
         .Columns(iCampo("Nº Conta")).Width = 60
         .Columns(iCampo("Data")).Width = 90
         .Columns(iCampo("Atendimento")).Width = 60
         .Columns(iCampo("Venda")).Width = 60
         .Columns(iCampo("Valor")).Width = 100
         .Columns(iCampo("Cliente")).Width = (.Width / Screen.TwipsPerPixelX) - 271
         
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
         Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.Sair, pIconId:=10007, pCaption:="Sair", pBeginGroup:=True)
         
         Dim n As Object
         For Each n In .Controls
            n.Category = "BAR_CONTAPR"
         Next
         
         .Position = 0
         .Visible = True
      End With
   End With
End Sub
Private Sub NovaFatura()
   Dim MyFat As TL_Fatura
   
   Set MyFat = New TL_Fatura
   With MyFat
      Set .Sys = mvarSys
      .Show vbModal
   End With
   Set MyFat = Nothing
End Sub
Private Sub EditarFatura(pIDCONTA As Long)
   Dim MyFat As TL_Fatura
   
   Set MyFat = New TL_Fatura
   With MyFat
      Set .Sys = mvarSys
      .IDLOJA = mvarIDLOJA
      .IDCONTA = pIDCONTA
      .Show vbModal
   End With
   MyFat = Nothing
   Exit Sub
Saida:
   ShowError
End Sub
