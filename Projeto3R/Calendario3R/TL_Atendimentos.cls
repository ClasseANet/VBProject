VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Atendimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe     As FrmAtendimentos
Attribute mvarMe.VB_VarHelpID = -1
Private Enum eIconId
   Novo = 1
   Editar
   'RPS
   'NF
   'ExpRPS
   'ImpNF
   Sair
End Enum

Dim iCampo As Collection

Private RsCad           As Object
Private mvarSys         As Object
Private mvarConsulta    As String
Private mvarIDLOJA      As Integer
Public Sub F_Refresh()
   Call PopulaGrdAtends(True)
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
   Set mvarMe = New FrmAtendimentos
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_CommandBarsExecute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   Select Case Control.Id
      'Case eIconId.Novo:   Call NovoAtends
      Case eIconId.Editar: Call EditarAtend(mvarMe.GrdAtends.SelectedRows(0).Record(iCampo("Nº Atend.")).Value)
      'Case eIconId.RPS:    'Call EditarRPS
      'Case eIconId.NF:     'Call EditarNF
      'Case eIconId.ExpRPS: Call ExportarRPS
      'Case eIconId.ImpNF:  Call ImportarNF
      Case eIconId.Sair:   Unload mvarMe
      Case "Sair":
   End Select
End Sub

Private Sub mvarMe_GrdAtendsBeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
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

Private Sub mvarMe_GrdAtendsKeyUp(KeyCode As Integer, Shift As Integer)
   Call mvarMe_KeyUp(KeyCode, Shift)
End Sub

Private Sub mvarMe_GrdAtendsRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   Call EditarAtend(mvarMe.GrdAtends.SelectedRows(0).Record(iCampo("Nº Atend.")).Value)
End Sub
Private Sub mvarMe_Activate()
'   Call MontarTela
'   Call PopulaTela
   
   mvarMe_Resize
   If GetTag(mvarMe, "1Vez", 0) = 0 Then
      Call SetTag(mvarMe, "1Vez", 1)
      Call PosicionaForm
   End If

'   Call MontarToolbar

   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF5 Then
      Call PopulaGrdAtends(pSaldo:=(Shift <> 0))
   ElseIf KeyCode = vbKeyEscape Then
         mvarMe.txtFiltrar.Text = mvarMe.txtFiltrar.Tag
         Call PopulaGrdAtends
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
      With .GrdAtends
         .BorderStyle = xtpBorderFrame
         .PaintManager.ColumnStyle = xtpColumnOffice2007
         .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
         .PaintManager.RefreshMetrics
         .Redraw
      End With
   End With
End Sub
Private Sub PopulaTela()
   Call PopulaGrdAtends
   mvarMe_Resize
End Sub
Private Sub mvarMe_Resize()
   On Error Resume Next
   'mvarMe.SccContato.Move 0, 0, mvarMe.ScaleWidth, 400

   mvarMe.GrdAtends.Left = 0
   'mvarMe.GrdAtends.Top = mvarMe.SccContato.Top + mvarMe.SccContato.Height
   mvarMe.GrdAtends.Top = 400
   mvarMe.GrdAtends.Width = mvarMe.ScaleWidth - 50
   mvarMe.GrdAtends.Height = mvarMe.ScaleHeight - mvarMe.GrdAtends.Top - 50

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
      mvarMe.GrdAtends.FilterText = mvarMe.txtFiltrar.Text
      mvarMe.GrdAtends.Populate
   End If
End Sub

Private Sub mvarMe_txtFiltrarLostFocus()
   
   If mvarMe.txtFiltrar.Text = "" Then mvarMe.txtFiltrar.Text = mvarMe.txtFiltrar.Tag
   mvarMe.txtFiltrar.ForeColor = 12632256
End Sub
Private Sub PopulaGrdAtends(Optional bRefresh As Boolean, Optional pSaldo As Boolean)
   Dim i As Long
   Dim nRow As Long
   Dim Sql  As String
   Dim bSaldo As Boolean
   bSaldo = pSaldo '(Mid(mvarMe.txtFiltrar.Text1, 2) = "-1")
   With mvarMe.GrdAtends
      .SetImageList mvarMe.ImageList1
     
      If bRefresh Then
         nRow = .SelectedRows(0).Index
      End If
      If mvarConsulta = "" Then

         Sql = "Select Right('000000'+Rtrim(Cast(A.IDATENDIMENTO As Char)),6) [Nº Atend.]" & vbNewLine
         Sql = Sql & ", Cast(Convert(Char(10),A.DTATEND, 103) As SmallDatetime) [Data]" & vbNewLine
         Sql = Sql & ", Convert(Char(5),A.HHINI, 108) [Inicio]" & vbNewLine
         Sql = Sql & ", Convert(Char(5),A.HHFIM, 108) [Fim]" & vbNewLine
         Sql = Sql & ", C.NOME [Cliente]" & vbNewLine
         Sql = Sql & ", S.CODSALA [Sala]" & vbNewLine
         Sql = Sql & ", M.CODMAQUINA [Máquina]" & vbNewLine
         Sql = Sql & ", F.NOME [Operadora], A.OBS [Obs.]" & vbNewLine
         
         If bSaldo Then
            'Sql = Sql & " , ("
            'Sql = Sql & "   (Select Count(SS.IDSESSAO) From OSESSAO SS Where SS.IDLOJA=A.IDLOJA And SS.IDATENDIMENTO=A.IDATENDIMENTO And SS.IDTPSERVICO<>1)"
            'Sql = Sql & " - (Select Count(SS2.IDSESSAO) From OSESSAO SS2 Join OATENDIMENTO_VENDA AV On AV.IDLOJA=SS2.IDLOJA And AV.IDATENDIMENTO=SS2.IDATENDIMENTO And AV.IDVENDA=SS2.IDVENDA Where SS2.IDLOJA=A.IDLOJA And SS2.IDATENDIMENTO=A.IDATENDIMENTO And SS2.IDTPSERVICO<>1)"
            'Sql = Sql & "    )[Saldo]"
         
         
            Sql = Sql & " , ("
            Sql = Sql & "   Select Count(SS.IDSESSAO) From OSESSAO SS Where SS.IDLOJA=A.IDLOJA And SS.IDATENDIMENTO=A.IDATENDIMENTO And SS.IDTPSERVICO<>1 And (isNull(SS.IDVENDA,0)=0 Or isNull(SS.IDITEM,0)=0)"
            Sql = Sql & "    )[Saldo]"
         Else
            Sql = Sql & ", '' [Saldo]"
         End If
         Sql = Sql & " From OATENDIMENTO A" & vbNewLine
         Sql = Sql & " Join OCLIENTE C On A.IDLOJA=C.IDLOJA And A.IDCLIENTE=C.IDCLIENTE" & vbNewLine
         Sql = Sql & " Join RFUNCIONARIO F On A.IDLOJA=F.IDLOJA And A.IDFUNCIONARIO=F.IDFUNCIONARIO " & vbNewLine
         Sql = Sql & " Left Join OSALA S On A.IDLOJA=S.IDLOJA And A.IDSALA=S.IDSALA " & vbNewLine
         Sql = Sql & " Left Join OMAQUINA M On A.IDLOJA=M.IDLOJA And A.IDMAQUINA=M.IDMAQUINA" & vbNewLine
         Sql = Sql & " Where A.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
         'Sql = Sql & " And V.SITVENDA='10'"
         'Sql = Sql & " And V.VLVENDA-V.VLDESC>0"
         If bSaldo Then
            Sql = Sql & " Order By Saldo Desc, Data Desc, Inicio Desc"
         Else
            Sql = Sql & " Order By Data Desc, Inicio Desc"
         End If
      Else
         Sql = mvarConsulta
      End If
            
      If mvarSys.xDb.AbreTabela(Sql, RsCad) Then
         Call FillRCFromRS(RsCad, mvarMe.GrdAtends, pCollColumn:=iCampo, pCurrency:=True)
      End If
      
      'For i = 0 To .Records.Count - 1
      '   .Records(i).Item(1).Icon = 0
      'Next
      
      If mvarMe.txtFiltrar.Text = "" Or mvarMe.txtFiltrar.Text = mvarMe.txtFiltrar.Tag Or mvarMe.txtFiltrar.Text = "-1" Then
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
      Call DefineGrdAlinha(bSaldo)
      
      
      .Redraw
   End With

End Sub
Private Sub DefineGrdAlinha(Optional pSaldo As Boolean)
   Dim vCol As ReportColumn
   If iCampo Is Nothing Then Exit Sub
   With mvarMe.GrdAtends
      For Each vCol In .Columns
         vCol.HeaderAlignment = xtpAlignmentCenter
         'vCol.Sortable = False
      Next
      If iCampo.Count <> 0 Then
         .Columns(iCampo("Nº Atend.")).HeaderAlignment = xtpAlignmentCenter
         .Columns(iCampo("Nº Atend.")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Data")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Inicio")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Fim")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Cliente")).HeaderAlignment = xtpAlignmentLeft
         .Columns(iCampo("Cliente")).Alignment = xtpAlignmentLeft
         .Columns(iCampo("Sala")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Máquina")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Operadora")).HeaderAlignment = xtpAlignmentLeft
         .Columns(iCampo("Operadora")).Alignment = xtpAlignmentLeft
         .Columns(iCampo("Obs.")).HeaderAlignment = xtpAlignmentLeft
         .Columns(iCampo("Obs.")).Alignment = xtpAlignmentLeft
         .Columns(iCampo("Saldo")).HeaderAlignment = xtpAlignmentCenter
         .Columns(iCampo("Saldo")).Alignment = xtpAlignmentCenter
         
         
         .AutoColumnSizing = False
         .Columns(iCampo("Nº Atend.")).Width = 60
         .Columns(iCampo("Data")).Width = 90
         .Columns(iCampo("Inicio")).Width = 45
         .Columns(iCampo("Fim")).Width = 45
         .Columns(iCampo("Cliente")).Width = 150
         .Columns(iCampo("Sala")).Width = 40
         .Columns(iCampo("Máquina")).Width = 60
         .Columns(iCampo("Operadora")).Width = 80
         .Columns(iCampo("Saldo")).Width = 40
         .Columns(iCampo("Saldo")).Visible = pSaldo
         .Columns(iCampo("Obs.")).Width = (.Width / Screen.TwipsPerPixelX) - 600 - IIf(pSaldo, .Columns(iCampo("Saldo")).Width, 0)
         
         '.Populate
      End If
   End With
End Sub

Private Sub MontarToolbar()
   Dim MyToolBar As CommandBar
   Dim oControl As CommandBarControl
   
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

         Set oControl = CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.Novo, pIconId:=10001, pCaption:="Novo")
         oControl.Enabled = False
         Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.Editar, pIconId:=10002, pCaption:="Editar")
         'Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.RPS, pIconId:=10003, pCaption:="RPS", pBeginGroup:=True)
         'Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.NF, pIconId:=10004, pCaption:="N.F.")
         'Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.ExpRPS, pIconId:=10005, pCaption:="Exportar RPS", pBeginGroup:=True)
         'Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.ImpNF, pIconId:=10006, pCaption:="Importar NF")
         Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.Sair, pIconId:=10007, pCaption:="Sair", pBeginGroup:=True)
         
         Dim n As Object
         For Each n In .Controls
            n.Category = "BAR_ATEND"
         Next
         
         .Position = 0
         .Visible = True
      End With
   End With
End Sub
Private Sub NovoAtends()
   Dim MyAtend As TL_Atendimento
   
   Set MyAtend = New TL_Atendimento
   With MyAtend
      Set .Sys = mvarSys
      .Show vbModal
   End With
   Set MyAtend = Nothing
End Sub
Private Sub EditarAtend(pIDATEND As Long)
   Dim MyAtend As TL_Atendimento
   
   Screen.MousePointer = vbHourglass
   
   Set MyAtend = New TL_Atendimento
   With MyAtend
      Set .Sys = mvarSys
      .IDLOJA = mvarIDLOJA
      .IDATENDIMENTO = pIDATEND
      .Show vbModal
   End With
   Screen.MousePointer = vbDefault
   Exit Sub
Saida:
   Screen.MousePointer = vbDefault
   ShowError
End Sub
