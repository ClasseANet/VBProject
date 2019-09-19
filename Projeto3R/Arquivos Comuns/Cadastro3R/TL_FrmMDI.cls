VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_FrmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe  As FrmMDI
Attribute mvarMe.VB_VarHelpID = -1
Dim mvarPane  As TL_ShortBar

Private mvarSys      As Object
Private mvarItemCad  As Integer
Private mvarIDMODU   As String
Private mvarTLOnFocus  As Object
Public Property Set TLOnFocus(ByVal vData As Object)
   Set mvarTLOnFocus = vData
End Property
Public Property Get TLOnFocus() As Object
   Set TLOnFocus = mvarTLOnFocus
End Property
Public Property Let ItemCad(ByVal vData As Integer)
   mvarItemCad = vData
End Property
Public Property Get ItemCad() As Integer
   ItemCad = mvarItemCad
End Property
Public Property Let IDMODU(ByVal vData As String)
   mvarIDMODU = vData
End Property
Public Property Get IDMODU() As String
   IDMODU = mvarIDMODU
End Property

Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Sub Show()
   '* Carregar Form
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmMDI
End Sub
Private Sub mvarMe_Activate()
   mvarMe_Resize
   If GetTag(mvarMe, "1Vez", 0) = 0 Then
      Call SetTag(mvarMe, "1Vez", 1)
      Call PosicionaForm
   End If
   If mvarMe.CmdEditar.Enabled Then
      mvarMe.CmdEditar.SetFocus
   End If
  
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmdEditarClick()
   Dim nRow As Long
   Screen.MousePointer = vbHourglass
   
   On Error GoTo Saida
   If Trim(mvarPane.TL_CAD) = "" Then
      'mvarPane.TL_CAD = "TL_CADGENERICO"
      ExibirInformacao "Função não disponível nesta versão."
      GoTo Fim
   End If
      
   If TypeName(mvarTLOnFocus) <> mvarPane.TL_CAD Then
      Set mvarTLOnFocus = CriarObjeto("CADASTRO3R." & mvarPane.TL_CAD, False)
   End If
   Set mvarTLOnFocus.Sys = mvarSys
   Set mvarTLOnFocus.ShortBar = mvarPane
   Set mvarTLOnFocus.RptRec = mvarPane.TLCadastro.FrmCad.GrdCadastro.SelectedRows(0)
   Screen.MousePointer = vbDefault
   mvarTLOnFocus.Show
   
   With mvarPane.TLCadastro.FrmCad.GrdCadastro
      If .SelectedRows.Count > 0 Then nRow = .SelectedRows(0).Index
   End With
   'mvarPane.TLCadastro.F_Refresh
   'Set mvarTLOnFocus = Nothing
   'mvarPane.TLCadastro.PopulaTela
   Call Reload
   With mvarPane.TLCadastro.FrmCad.GrdCadastro
      If .SelectedRows.Count > 0 Then .Rows(.SelectedRows(0).Index).Selected = False
      .Rows(nRow).Selected = True
      .Redraw
   End With
   GoTo Fim
Saida:
   ExibirInformacao "Função não disponível nesta versão."
Fim:
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmdExcluirClick()
'   If Trim(mvarPane.TL_CAD) = "" Then
      'ExibirInformacao "Função não disponível nesta versão."
'      ExibirInformacao "Função não pode ser realizada desta tela." & vbNewLine & "Favor utilizar tela de cadastro."
'      Exit Sub
'   End If
   Screen.MousePointer = vbHourglass
   Set mvarTLOnFocus = CriarObjeto("CADASTRO3R." & mvarPane.TL_CAD, False)
   If mvarTLOnFocus Is Nothing Then
      ExibirInformacao "Função não pode ser realizada desta tela." & vbNewLine & "Favor utilizar tela de cadastro."
   Else
      Set mvarTLOnFocus.Sys = mvarSys
      Set mvarTLOnFocus.RptRec = mvarPane.TLCadastro.FrmCad.GrdCadastro.SelectedRows(0)
      If mvarTLOnFocus.Excluir Then
         mvarPane.TLCadastro.F_Refresh
      End If
      Set mvarTLOnFocus = Nothing
   End If
   If mvarMe.CmdEditar.Enabled Then
      mvarMe.CmdEditar.SetFocus
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmdNovoClick()
   On Error GoTo Saida
   
   Screen.MousePointer = vbHourglass
   
   If Trim(mvarPane.TL_CAD) = "" Then
      ExibirInformacao "Função não disponível nesta versão."
      GoTo Fim
   End If
   Set mvarTLOnFocus = CriarObjeto("CADASTRO3R." & mvarPane.TL_CAD, False)
   Set mvarTLOnFocus.Sys = mvarSys
   Set mvarTLOnFocus.ShortBar = mvarPane
   Set mvarTLOnFocus.RptRec = Nothing
   Screen.MousePointer = vbDefault
   mvarTLOnFocus.Show
   'Set mvarTLOnFocus = Nothing
   Call Reload
   If mvarMe.CmdEditar.Enabled Then
      mvarMe.CmdEditar.SetFocus
   End If
   
   GoTo Fim
Saida:
   ExibirInformacao "Função não disponível nesta versão."
Fim:
   Screen.MousePointer = vbDefault
End Sub
Private Sub Reload()
   Dim xItem As Object
   If True Then
      mvarMe_Load
      mvarMe_Resize
   Else
   With mvarPane.frm
      If .wndTaskPanel.FocusedItem Is Nothing Then
         Set xItem = .wndTaskPanel.Find(mvarItemCad)
         If xItem Is Nothing Then
            Call mvarPane.ItemClick(.wndTaskPanel.Groups(1).Items(1), True)
         Else
            Call mvarPane.ItemClick(xItem, True)
         End If
      Else
         Call mvarPane.ItemClick(.wndTaskPanel.FocusedItem, True)
      End If
   End With
   End If
End Sub
Private Sub mvarMe_CmdSairClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_Load()
   'Call SetMDI(mvarMe.hWnd, mvarSys.MDI.hWnd)
   Set mvarSys.ObjAux = mvarMe
   If Not mvarPane Is Nothing Then
      mvarItemCad = mvarPane.ItemCad
   End If
   Set mvarPane = New TL_ShortBar
   With mvarPane
      Set .Sys = mvarSys
      Set .MDI = mvarMe
      .ItemCad = mvarItemCad
      .Show
   End With
   Call PosicionaForm
   Screen.MousePointer = vbDefault
End Sub
Private Sub PosicionaForm()
   Dim Left As Long
   Dim Top As Long
   Dim Right As Long
   Dim Bottom As Long
   On Error Resume Next
   mvarSys.MDI.CommandBars.GetClientRect Left, Top, Right, Bottom
   mvarSys.MDI.CommandBars(1).GetWindowRect Left, Top, Right, Bottom
      
   'mvarMe.Left = mvarSys.MDI.Left + 60
   mvarMe.Top = Bottom
   'mvarMe.Width = mvarSys.MDI.Width - mvarMe.Left - 120
   'mvarMe.Height = mvarSys.MDI.Height - mvarMe.Top - 420
   
   
   mvarMe.Left = mvarSys.MDI.Left + 60
   'mvarMe.Top = mvarSys.MDI.Top + 780
   mvarMe.Width = mvarSys.MDI.Width - mvarMe.Left - 120
   mvarMe.Height = mvarSys.MDI.Height - mvarMe.Top - 420
End Sub

Private Sub mvarMe_Resize()
   mvarMe.PictBottom.BorderStyle = 0
   mvarMe.PictBottom.Height = 675
   
   mvarMe.GrpBoxBottom.Top = 0
   mvarMe.GrpBoxBottom.Left = 0
   mvarMe.GrpBoxBottom.Height = mvarMe.PictBottom.Height
   mvarMe.GrpBoxBottom.Width = mvarMe.PictBottom.Width
   
   mvarMe.TabPgBotton.Visible = True
   mvarMe.TabPgBotton.Top = 120
   mvarMe.TabPgBotton.Left = 0
   mvarMe.TabPgBotton.Height = mvarMe.PictBottom.Height
   mvarMe.TabPgBotton.Width = mvarMe.PictBottom.Width
   
   mvarMe.CmdSair.Left = (mvarMe.Width / 2) + ((mvarMe.CmdSair.Width) * 4 / 2)  ' 10000
   mvarMe.CmdEditar.Left = mvarMe.CmdSair.Left - mvarMe.CmdSair.Width - 120
   mvarMe.CmdExcluir.Left = mvarMe.CmdEditar.Left - mvarMe.CmdEditar.Width - 120
   mvarMe.CmdNovo.Left = mvarMe.CmdExcluir.Left - mvarMe.CmdExcluir.Width - 270
End Sub
