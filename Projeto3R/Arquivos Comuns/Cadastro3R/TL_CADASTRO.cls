VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const sPesquisa = "Pesquisar..."
Private WithEvents mvarMe     As FrmCadastro
Attribute mvarMe.VB_VarHelpID = -1

Dim iCampo As Collection

Private RsCad           As Object
Private mvarSys         As Object
Private mvarConsulta    As String
Private mvarTitulo      As String
Private mvarTL_CAD      As String
Private mvaroTLCAD      As Object
Private mvarTAGCAMPOS   As String
Private mvarIDMODU      As String
Private mvarMDI         As Object
Public Sub F_Refresh()
   Call PopulaGrdCadastro(True)
End Sub
Public Property Get FrmCad() As Object 'FrmCadastro
   Set FrmCad = mvarMe
End Property
Public Property Set MDI(ByVal vData As Object)
   Set mvarMDI = vData
End Property
Public Property Get MDI() As Object
   Set MDI = mvarMDI
End Property

Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
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
Public Property Get Titulo() As String
   Titulo = mvarTitulo
End Property
Public Property Let Titulo(ByVal vData As String)
   mvarTitulo = vData
End Property
Public Property Get TL_CAD() As String
   TL_CAD = mvarTL_CAD
End Property
Public Property Set oTLCAD(ByVal vData As Object)
   Set mvaroTLCAD = vData
End Property
Public Property Get oTLCAD() As Object
   Set oTLCAD = mvaroTLCAD
End Property
Public Property Let TL_CAD(ByVal vData As String)
   mvarTL_CAD = vData
End Property
Public Property Get TAGCAMPOS() As String
   TAGCAMPOS = mvarTL_CAD
End Property
Public Property Let TAGCAMPOS(ByVal vData As String)
   mvarTAGCAMPOS = vData
End Property
Public Property Get IDMODU() As String
   IDMODU = mvarIDMODU
End Property
Public Property Let IDMODU(ByVal vData As String)
   mvarIDMODU = vData
End Property
Public Sub Show(Optional pReload)
   If pReload Then
      Call PopulaGrdCadastro
   Else
      Load mvarMe
      '* Acoplar Form no MDI
      Call AcoplarForm(mvarMe, 2, mvarSys, True, mvarMDI)
      'mvarMe.Show
      
      '* Exibir Menu no MDI
      'With mvarSys.MDI.CommandBars
      '   If Not .FindControl(, Mnu_Contato, False) Is Nothing Then
      '      .FindControl(, Mnu_Contato).Visible = True
      '   End If
      'End With
   End If
   '* Definir foco no Lista de Clientes
   With mvarMe.GrdCadastro
      If .Visible And .Enabled Then
         .SetFocus
      End If
   End With
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmCadastro
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_GrdCadastroRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   On Error GoTo Saida
   DoEvents
   Screen.MousePointer = vbHourglass
   If Not mvarMDI Is Nothing Then
      mvarMDI.Editar
   End If
   GoTo Fim
Saida:
   ShowError
Fim:
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_Activate()
   'Call PopulaTela
   mvarMe.GrdCadastro.Height = mvarMe.Height - 1075
End Sub
Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF5 Then
      Call PopulaGrdCadastro
   ElseIf KeyCode = vbKeyInsert Then
      If Not mvarMDI Is Nothing Then mvarMDI.Novo
   ElseIf KeyCode = vbKeyReturn Then 'Or KeyCode = vbKeyMenu Then
      If Not mvarMDI Is Nothing Then mvarMDI.Editar
   ElseIf KeyCode = vbKeyDelete Then
      If Not mvarMDI Is Nothing Then mvarMDI.Excluir
   
   ElseIf KeyCode = vbKeyEscape Then
         mvarMe.txtFiltrar.Text = sPesquisa
         Call PopulaGrdCadastro
   ElseIf (Asc(UCase(Chr(KeyCode))) >= Asc("A") And Asc(UCase(Chr(KeyCode))) >= Asc("Z")) Or (Asc(UCase(Chr(KeyCode))) >= Asc("0") And Asc(UCase(Chr(KeyCode))) >= Asc("9")) Then
      If Not mvarMe.ActiveControl Is mvarMe.txtFiltrar Then
         mvarMe.txtFiltrar.Text = UCase(Chr(KeyCode))
         mvarMe.txtFiltrar.SetFocus
         mvarMe.txtFiltrar.SelStart = Len(mvarMe.txtFiltrar.Text)
      End If
   End If
End Sub
Private Sub mvarMe_Load()
   Call MontaTela
   Call PopulaTela
End Sub
Private Sub MontaTela()
   'ReportControlGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   With mvarMe
      With .SccContato
         .VisualTheme = xtpShortcutThemeOffice2003
         .GradientHorizontal = False
         .Caption = mvarTitulo
      End With
      With .GrdCadastro
         .BorderStyle = xtpBorderFrame
         .PaintManager.ColumnStyle = xtpColumnOffice2007
         .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
         .PaintManager.RefreshMetrics
         .Redraw
      End With
      mvarMe.GrdCadastro.Height = .Height - 1075
   End With
End Sub
Public Sub PopulaTela()
   Call PopulaGrdCadastro
   mvarMe_Rezise
End Sub
Private Sub mvarMe_Rezise()
   On Error Resume Next
   mvarMe.SccContato.Move 0, 0, mvarMe.ScaleWidth, 400

   mvarMe.GrdCadastro.Left = 0
   mvarMe.GrdCadastro.Top = mvarMe.SccContato.Top + mvarMe.SccContato.Height
   mvarMe.GrdCadastro.Width = mvarMe.ScaleWidth - 50
   mvarMe.GrdCadastro.Height = mvarMe.ScaleHeight - mvarMe.GrdCadastro.Top - 675
   'mvarMe.GrdCadastro.Height = mvarMe.Height - 1075

   mvarMe.txtFiltrar.Left = mvarMe.Width - 2800
   mvarMe.imgLupa.Left = mvarMe.Width - 680

   mvarMe.txtFiltrar.Top = (mvarMe.SccContato.Height - mvarMe.txtFiltrar.Height) / 2
   mvarMe.imgLupa.Top = mvarMe.txtFiltrar.Top + 5
   mvarMe.imgLupa.Height = mvarMe.txtFiltrar.Height - 30
   
   mvarMe.txtFiltrar.ZOrder 0
   mvarMe.imgLupa.ZOrder 0
End Sub

Private Sub mvarMe_txtFiltrarGotFocus()
   mvarMe.txtFiltrar.ForeColor = vbBlack
   
   If mvarMe.txtFiltrar.Text = sPesquisa Then
      mvarMe.txtFiltrar.Text = ""
   End If
End Sub

Private Sub mvarMe_txtFiltrarKeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      mvarMe.GrdCadastro.FilterText = mvarMe.txtFiltrar.Text
      mvarMe.GrdCadastro.Populate
   End If
End Sub

Private Sub mvarMe_txtFiltrarLostFocus()
   
   If mvarMe.txtFiltrar.Text = "" Then mvarMe.txtFiltrar.Text = sPesquisa
   mvarMe.txtFiltrar.ForeColor = 12632256
End Sub

Private Sub mvarMe_Unload(Cancel As Integer)
  On Error Resume Next
'   Unload FrmLembrete
End Sub
'=============================
'=============================
Private Sub PopulaGrdCadastro(Optional bRefresh As Boolean)
   Dim i As Long

   Dim sTagCampos As String
   Dim sTag As String
   Dim nRow As Long
      
   On Error Resume Next
   With mvarMe.GrdCadastro
      '.SetImageList mvarMe.ImageList1
     
      If bRefresh Then
         nRow = .SelectedRows(0).Index
      End If
      If mvarSys.Xdb.AbreTabela(mvarConsulta, RsCad) Then
         Call FillRCFromRS(RsCad, mvarMe.GrdCadastro, pCollColumn:=iCampo)
      End If
      
      'For i = 0 To .Records.Count - 1
      '   .Records(i).Item(1).Icon = 0
      'Next
      
      If mvarMe.txtFiltrar.Text = "" Or mvarMe.txtFiltrar.Text = sPesquisa Then
         .FilterText = ""
      Else
         .FilterText = mvarMe.txtFiltrar.Text
      End If
      
      sTagCampos = mvarTAGCAMPOS
      On Error Resume Next
      While InStr(sTagCampos, ":") <> 0
         sTag = Mid(sTagCampos, 1, InStr(sTagCampos, ":") - 1)
         mvarMe.GrdCadastro.Columns(iCampo(Trim(GetTag(sTag, "CAMPO")))).Visible = (GetTag(sTag, "VISIBLE") = "S")
         sTagCampos = Mid(sTagCampos, InStr(sTagCampos, ":") + 1)
      Wend
      
      
      .Populate
      If bRefresh Then
         If nRow > .Rows.Count - 1 Then
            nRow = .Rows.Count - 1
         End If
         .Rows(nRow).EnsureVisible
         .Rows(nRow).Selected = True
      End If
      .Redraw
   End With

End Sub

