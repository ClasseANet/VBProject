VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#11.2#0"; "Codejock.CommandBars.v11.2.2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#11.2#0"; "Codejock.DockingPane.v11.2.2.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#11.2#0"; "Codejock.SkinFramework.v11.2.2.ocx"
Begin VB.MDIForm MDI 
   BackColor       =   &H8000000C&
   Caption         =   "Projeto 3R - Módulo Loja"
   ClientHeight    =   4350
   ClientLeft      =   3510
   ClientTop       =   4935
   ClientWidth     =   10455
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PctMDI 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      Height          =   2370
      Left            =   0
      ScaleHeight     =   2310
      ScaleWidth      =   10395
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   10455
      Begin VB.PictureBox PctMenuCOLIGADA 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   360
         ScaleHeight     =   345
         ScaleWidth      =   10455
         TabIndex        =   5
         Top             =   1560
         Visible         =   0   'False
         Width           =   10455
         Begin VB.PictureBox PctSaida 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   10140
            MousePointer    =   99  'Custom
            Picture         =   "MDI.frx":1582
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   6
            Top             =   60
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image ImgAlias 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   8565
            Picture         =   "MDI.frx":23C4
            Top             =   90
            Width           =   240
         End
         Begin VB.Label LblColigada 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Sem Coligada!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   300
            Left            =   6720
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   30
            Width           =   1755
         End
         Begin VB.Image ImgMenuCOLIGADA 
            Height          =   345
            Left            =   240
            Picture         =   "MDI.frx":294E
            Stretch         =   -1  'True
            Top             =   0
            Width           =   615
         End
      End
   End
   Begin VB.PictureBox picHolder 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   10455
      TabIndex        =   0
      Top             =   3810
      Width           =   10455
      Begin VB.Timer Timer 
         Interval        =   500
         Left            =   2640
         Top             =   0
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBar 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1815
         _Version        =   720898
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   93
         Text            =   "Loading..."
         ForeColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Scrolling       =   1
         Appearance      =   4
         FlatStyle       =   -1  'True
         BarColor        =   -2147483636
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   7920
         TabIndex        =   8
         Top             =   120
         Width           =   1815
         _Version        =   720898
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   93
         Text            =   "Loading..."
         ForeColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Scrolling       =   2
         Appearance      =   4
         UseVisualStyle  =   -1  'True
         FlatStyle       =   -1  'True
         BarColor        =   -2147483636
      End
      Begin XtremeSuiteControls.Label LblPercentual 
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   120
         Width           =   585
         _Version        =   720898
         _ExtentX        =   1032
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "100%"
         ForeColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   3720
         TabIndex        =   2
         Top             =   120
         Width           =   1815
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   3480
      Top             =   2760
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2280
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":323B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   1560
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin XtremeSuiteControls.PopupControl PopupControl 
      Index           =   0
      Left            =   3000
      Top             =   2640
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   960
      Top             =   2640
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane DockingPaneManager 
      Left            =   600
      Top             =   2640
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgMenu 
      Left            =   240
      Top             =   2640
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "MDI.frx":3F15
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FComando As String
Public WithEvents StatusBar As XtremeCommandBars.StatusBar
Attribute StatusBar.VB_VarHelpID = -1
Public WithEvents Workspace As TabWorkspace
Attribute Workspace.VB_VarHelpID = -1

Private nStatusButton As Integer
Private Const nIndiceColigada = 90000
Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   Dim sIDMODU    As String
   Dim sVBSCRIPT  As String
   
   On Error GoTo Trata_Erro
   DoEvents
   Screen.MousePointer = vbHourglass
     
   sVBSCRIPT = GetTag(Control.Parameter, "VBSCRIPT", "")
   If sVBSCRIPT = "" Then
      sIDMODU = GetTag(Control.Parameter, "IDMODU", "")
      Select Case sIDMODU
         Case "ADMEND":
             End
         Case "SECAOSIS_IDCOLIGADA":
            '* Trocar ID da Coligada
            Sys.IDCOLIGADA = Control.Id
            Me.LblColigada.Caption = Control.Caption
               
            If MDI.CommandBars.Count >= 2 Then
               MDI.CommandBars(2).Controls(1).Caption = Control.Caption
               MDI.CommandBars(2).Controls(1).ToolTipText = Control.Caption
            End If
               
         Case "SECAOSIS_IDPROJ":
            If MDI.CommandBars.Count >= 2 Then
               If MDI.CommandBars(2).Controls.Count >= 3 Then
                  MDI.CommandBars(2).Controls(3).Caption = Control.Caption
               End If
            End If
      End Select
   End If
      
   GoTo Saida
Trata_Erro:
   If Err = 0 Then
      Call ExibirAviso("Script não cadastrado", Sys.CODSIS)
   Else
      Call ExibirAviso("Script não cadastrado" & vbNewLine & vbNewLine & Err & "-" & Error, Sys.CODSIS)
   End If
Saida:
   Screen.MousePointer = vbDefault
End Sub


Private Sub DockingPaneManager_Resize()
   Dim PaneACx As Integer
   Dim PaneACy As Integer
   Dim PaneBCx As Integer
   Dim PaneBCy As Integer
   Dim PaneCCx As Integer
   Dim PaneCCy As Integer
PaneACx = PaneACx
End Sub

Private Sub DockingPaneManager_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
   Dim PaneACx As Integer
   Dim PaneACy As Integer
   Dim PaneBCx As Integer
   Dim PaneBCy As Integer
   Dim PaneCCx As Integer
   Dim PaneCCy As Integer

PaneACx = PaneACx
End Sub

Private Sub ImgAlias_Click()
   Dim oRS           As Object
   Dim Sql           As String
   Dim oCommBar      As CommandBarControl
   Dim Popup         As CommandBar
   Dim ContextMenu   As CommandBar
   
   Set ContextMenu = CommandBars.ContextMenus.Add(0, "Context Menu")
   
   Sql = "Select C.IDCOLIGADA, C.NMCOLIGADA"
   Sql = Sql & " From COLIGADA C, USUARIO_COLIGADA U"
   Sql = Sql & " Where C.IDCOLIGADA=U.IDCOLIGADA"
   Sql = Sql & " And U.IDUSU=" & Sys.xdb.SqlStr(Sys.IDUSU)
   
   If Sys.xdb.AbreTabela(Sql, oRS) Then
      While Not oRS.EOF
         Set oCommBar = ContextMenu.Controls.Add(xtpControlButton, oRS("IDCOLIGADA"), oRS("NMCOLIGADA"))
         oCommBar.Parameter = SetTag(oCommBar.Parameter, "IDMODU", "SECAOSIS_IDCOLIGADA")
         oCommBar.Parameter = SetTag(oCommBar.Parameter, "CARREGADO", "1")
         oRS.MoveNext
      Wend
      Set oRS = Nothing
      
      Set Popup = CommandBars.ContextMenus.Find(0)
      Popup.ShowPopup
      CommandBars.ContextMenus.DeleteAll
      
   Else
   
   End If
End Sub
Private Sub ImgMenuCOLIGADA_DblClick()
   PctMenuCOLIGADA_DblClick
End Sub
Public Sub Reload(Optional bFull As Boolean = True)
   gReload = True
   Dim StPane As StatusBarPane
   Screen.MousePointer = vbHourglass
   
   Call MDIForm_Initialize
   Call ConfigurarAmbiente(True)
   
   MontarMenu
   MontarToolbar
   'MontarStatusBar
   Set StatusBar = Me.CommandBars.StatusBar
   StatusBar.Pane(1).Text = Sys.USER.IDUSU
   StatusBar.Pane(0).Text = "[" & Sys.xdb.Alias & "]"

   If bFull Then
      'MontarPanes
      LoadMenuDefault
   End If
   Call MDIForm_Resize
   
   gReload = False
   Screen.MousePointer = vbDefault
End Sub

Private Sub MDIForm_Activate()
'   me.DockingPaneManager.Panes(2).SetHandle 0
   If GetTag(Me, "PrimeiraVez", "S") = "S" Then
      Call SetTag(Me, "PrimeiraVez", "N")

     ' MontarMenu
     ' MontarToolbar
     ' MontarStatusBar
     ' MontarPanes
      
      If Not gDebug Then gDebug = gDebugGotoShow
      LoadMenuDefault
   End If
   If Sys.Propriedades("FCOMANDO") <> "" Then
      If Sys.Propriedades("FCOMANDO") = "End" Then
         End
      End If
   End If
   Sys.Propriedades("Debug") = gDebug
End Sub

Private Sub MDIForm_DblClick()
   Dim MyObj  As Object
   
   Set MyObj = CriarObjeto("SHORTBAR.TL_SHORTBAR")
   Set MyObj.Sys = Sys
   MyObj.Show
End Sub

Private Sub MDIForm_Initialize()
   On Error GoTo TrataErro
   If Not gReload Then
      Call InitCommonControls
   End If
   Set Sys.MDI = Me
   Me.PctMDI.Height = Me.Height
   Exit Sub
TrataErro:
   MsgBox Err & " - " & Error, vbOKOnly + vbCritical, "Atenção!"
End Sub
Private Sub MDIForm_Load()
   Screen.MousePointer = vbDefault
   
   ConfigurarAmbiente
   
   MontarMenu
   MontarToolbar
   MontarStatusBar
   MontarPanes
End Sub
Private Sub ConfigurarAmbiente(Optional bReload As Boolean)
   Dim sPath   As String
   Dim sArq    As String
   Dim sEstilo As String
   
   If Sys.CODSIS = "P3R" Then
      Sys.Style = 0
      Sys.Skin = IIf(Sys.xdb.Alias = "PRODUCAO", 2, 3)
      Sys.ShowTabWorkspace = True
   End If
   
   sPath = GetSpecialFolder(38) & "ClasseA\Arquivos Comuns\Styles\"
   'sPath = Environ("programfiles") & "\ClasseA\Arquivos Comuns\Styles\"
   Select Case Sys.Style
      Case 0: sArq = sPath & "WinXP.Luna.cjstyles"
      Case 1: sArq = sPath & "WinXP.Royale.cjstyles"
      Case 2: sArq = sPath & "Vista.cjstyles"
      Case 3: sArq = sPath & "Office2007.cjstyles"
   End Select
      
   If Not ExisteArquivo(sArq) Then Call ExtractResData("WINXPLUNA", "STYLE", sArq)
      
   Select Case Sys.Skin
      Case 0: sEstilo = ""                      '* SKIN ROYALE
      Case 1: sEstilo = "NormalBlue.ini"        '* SKIN LUNA BLUE
      Case 2: sEstilo = "NormalHomestead.ini"   '* SKIN LUNA OLIVE
      Case 3: sEstilo = "NormalMetallic.ini"    '* SKIN LUNA METALLIC
   End Select
    
   
   Set Me.Icon = MyLoadPicture
  
'MsgBox "sArq= " & sArq & vbNewLine & "sEstilo= " & sEstilo
   With Me.SkinFramework
      .LoadSkin sArq, sEstilo
      If Not bReload Then
         .ApplyWindow Me.hwnd
         .ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
      End If
   End With
   
   sArq = App.Path & "\Close.bmp"
   If Not ExisteArquivo(sArq) Then Call ExtractResData("CLOSE", "BMP", sArq)

   Me.picHolder.Height = 0
   
   'CommandBarsGlobalSettings.App = App
   With Me.CommandBars
      If Not bReload Then
         .GlobalSettings.App = App
      End If
     Set Workspace = .ShowTabWorkspace(Sys.ShowTabWorkspace)
      .VisualTheme = xtpThemeNativeWinXP   'xtpThemeWhidbey 'xtpThemeOffice2003
      .ToolTipContext.Style = xtpToolTipLuna
   End With
   
   If Workspace Is Nothing Then
      Set Workspace = Me.CommandBars.ShowTabWorkspace(Sys.ShowTabWorkspace)
   End If

   DockingPaneManager.SetCommandBars Me.CommandBars
End Sub
Private Sub CreateActions()
   With Me.CommandBars
      .EnableActions
   
      .Actions.Add 724, "", "", "", ""
   End With
End Sub
Private Sub MontarMenu()
   Dim Erro429    As Boolean
   Dim SetErro    As String

   On Error GoTo TrataErro

   If gDebug Then MsgBox "Menu is Nothing = " & IIf(MdiMenu Is Nothing, "True", "False")
   
   If MdiMenu Is Nothing Then
      If gDebug Then MsgBox "Set Menu = CriarObjeto(Menu.ControlMenu)"
      Set MdiMenu = CriarObjeto("Menu.ControlMenu")
   End If
   With MdiMenu
      Set .Sys = Sys
      .SisDebug = gDebug
      If gDebug Then MsgBox "Menu.ControlMenu.MontarMenu"
      Call .MontarMenu
   End With
   
   GoTo Saida
TrataErro:
   If Err = 429 Then
      MsgBox "MontarMenu" & vbNewLine & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "Atenção!"
   End If
Saida:
   Screen.MousePointer = vbDefault
End Sub
Private Sub LoadMenuDefault()
   If gDebug Then MsgBox "LoadMenuDefault"
   On Error Resume Next
         
   If True Then
      If Not MdiMenu Is Nothing Then
         Call MdiMenu.LoadMenuDefault
      End If
   Else
      Dim MyObj As Object
      Set MyObj = CreateObject("SHORTBAR3R.TL_SHORTBAR")
      Set MyObj.Sys = Sys
      MyObj.Show
   End If

   If gDebug Then MsgBox "LoadMenuDefault End"
End Sub
Private Sub MontarToolbar()
   'Me.CommandBars.FindControl(Type, Id, Visible, Recursive)
'   Call MontarToolbarDinamico(Me)
End Sub
Private Sub MontarStatusBar()
'   Dim StatusBar As XtremeCommandBars.IStatusBar
   Dim StPane As StatusBarPane
   Dim dData  As Date
'   If Not Me.Visible Then Exit Sub
   
   With Me.ProgressBar
      .Visible = True
      .UseVisualStyle = False
      .FlatStyle = Not .UseVisualStyle
      .BackColor = &H8000000F
      .BarColor = &H8000000C
      .Font.Bold = True
      .Text = ""
       
   End With
   
   Set StatusBar = Me.CommandBars.StatusBar
   With StatusBar
      .RemoveAll
      .Visible = True
      Set StPane = .AddPane(101)  '* CONEXÃO
      With StPane
         .Style = 0
         .Text = "[" & Sys.xdb.Alias & "]"
         .Alignment = xtpAlignmentCenter
         .TextColor = IIf(Sys.xdb.Conectado, vbBlack, &H8000000C)
         .ToolTip = "[" & Sys.xdb.Server & "].[" & Sys.xdb.dbName & "]"
         .Width = 100 ' Len(.Text) * 8
      End With
      
      Set StPane = .AddPane(102)  '* USUARIO
      With StPane
         .Style = 0
         .Text = Sys.USER.IDUSU
         .Alignment = xtpAlignmentCenter
         .TextColor = vbBlack
         .ToolTip = Sys.USER.NMUSU
         .Width = Len(.Text) * 8 '* (Screen.Width / Me.ScaleWidth)
      End With
      
      Set StPane = .AddPane(103)  '* PERFIL/EQUIPE...
      With StPane
         .Style = 0
         .Text = "equipe"
         .Alignment = xtpAlignmentCenter
         .TextColor = vbBlack
         .ToolTip = "Equipe"
         .Width = Len(.Text) * 8
      End With
      
      Set StPane = .AddPane(104)   '* STRETCH
      With StPane
         .Style = SBPS_STRETCH
         .Text = ""
         .Alignment = xtpAlignmentCenter
         .TextColor = vbBlack
         .ToolTip = ""
         '.Width = Len(.Pane(0).Text) * (8 * (Screen.Width / Me.ScaleWidth))
         .Handle = Me.ProgressBar.hwnd
      End With
      
      
      .AddPane 59137 'ID_INDICATOR_CAPS
      .AddPane 59138 'ID_INDICATOR_NUM
      .AddPane 59139 'ID_INDICATOR_SCRL
      
      If Sys.xdb.Conectado Then
         dData = Sys.xdb.Sysdate(3)
      Else
         dData = Now()
      End If
      Set StPane = .AddPane(108)  '* DATA
      
      With StPane
         .Style = 0
         .Text = Format(dData, "dd/mm/yyyy")
         .Alignment = xtpAlignmentCenter
         .TextColor = vbBlack
         .ToolTip = "Data do Servidor"
         .Width = Len(.Text) * 8
      End With
      Set StPane = .AddPane(109)  '* DATA
      With StPane
         .Style = 0
         .Text = Format(dData, "hh:mm")
         .Alignment = xtpAlignmentCenter
         .TextColor = vbBlack
         .ToolTip = "Hora do Servidor"
         .Width = Len(.Text) * 8
      End With
   End With
End Sub
Private Sub MontarPanes()
   Dim xPane As Pane
   Dim A As Pane
   Dim B As Pane
   Dim C As Pane
   Dim gPanes  As Integer
   Dim PaneACx As Integer
   Dim PaneACy As Integer
   Dim PaneBCx As Integer
   Dim PaneBCy As Integer
   Dim PaneCCx As Integer
   Dim PaneCCy As Integer
   
   gPanes = 2
   PaneACx = ReadIniFile(Sys.LocalReg, "Format", "PaneACx", 160)
   PaneACy = ReadIniFile(Sys.LocalReg, "Format", "PaneACy", 120)
   PaneBCx = ReadIniFile(Sys.LocalReg, "Format", "PaneBCx", 700)
   PaneBCy = ReadIniFile(Sys.LocalReg, "Format", "PaneBCy", 400)
   PaneCCx = ReadIniFile(Sys.LocalReg, "Format", "PaneCCx", 400)
   PaneCCy = ReadIniFile(Sys.LocalReg, "Format", "PaneCCy", 100)
   
   With Me.DockingPaneManager
      .DestroyAll
      If gPanes = 1 Then
         Set A = .CreatePane(1, PaneACx, PaneACy, DockTopOf)
         A.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable

      ElseIf gPanes = 2 Then
            Set A = .CreatePane(1, PaneACx, PaneACy, DockLeftOf, Nothing)
            A.Tag = 1
            A.TabColor = vbRed
            A.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable
            
            Set B = .CreatePane(2, PaneBCx, PaneBCy, DockRightOf, A)
            B.Tag = 2
            B.TabColor = vbBlue
            B.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable
            
      ElseIf gPanes = 3 Then
         Set A = .CreatePane(1, PaneACx, PaneACy, DockLeftOf, Nothing)
         A.Tag = 1
         
         Set B = .CreatePane(2, PaneBCx, PaneBCy, DockRightOf, A)
         B.Tag = 2
         
         Set C = .CreatePane(3, PaneCCx, PaneCCy, DockBottomOf, B)
         C.Tag = 3
         
      ElseIf gPanes = 4 Then
      
      End If
      .Options.HideClient = True
      .PaintManager.ShowCaption = False
      
   End With
   Me.CommandBars.RecalcLayout
End Sub
Private Sub MDIForm_Resize()
   If Me.Height >= 480 Then
      Me.PctMDI.Height = Me.Height - 480
   End If
   Me.PctMenuCOLIGADA.Width = Me.Width
   Me.ImgMenuCOLIGADA.Left = 0
   Me.ImgMenuCOLIGADA.Width = Me.PctMenuCOLIGADA.Width
   Me.ImgMenuCOLIGADA.ZOrder 1
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   Dim i As Integer
   
   On Error Resume Next
   For i = PopupControl.lbound To PopupControl.Ubound
      Call PopupControl(i).Close
   Next
   
   Call Sys.xdb.Executa("Delete GLOGIN Where IDUSU = " & SqlNum(gIDUSU))
End Sub

Private Sub PctMenuCOLIGADA_DblClick()
   Dim bVisible As Boolean
   Dim bExiste  As Boolean
   
   If Sys.USER.isSystem Then
      bExiste = Me.CommandBars.Count >= 2
      If bExiste Then bVisible = Me.CommandBars(2).Visible
      
      Me.PctMenuCOLIGADA.Top = IIf(bExiste And bVisible, 340, 0)
      
      If bExiste Then Me.CommandBars(2).Visible = Not bVisible
  End If
End Sub

Private Sub PopupControl_ItemClick(Index As Integer, ByVal Item As XtremeSuiteControls.IPopupControlItem)
   If Item.Id = 1 Then
      Call PopupControl(Index).Close
      If Index > 0 Then
         Unload PopupControl(Index)
      End If
   End If
End Sub

Private Sub StatusBar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   nStatusButton = Button
   Call SetTag(Me, "nStatusButton", Button)
End Sub

Private Sub StatusBar_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   On Error Resume Next
   Select Case nStatusButton
      Case 1 'Left
      Case 2 'Right
         Select Case StatusBar.HitTest(x, y).Id
           Case 101:   PaneConexao_Click    '* CONEXÃO
'           Case 102:   PaneUsusario_Click   '* USUARIO
           Case 103:   PaneEquipe_Click     '* PERFIL/EQUIPE...
'           Case 104:   PaneStrech_Click     '* STRETCH
'           Case 108:   PaneData_Click       '* DATA
'           Case 109:   PaneHora_Click       '* HORA
           
'           Case 59137: PaneCapsLock_Click   '* ID_INDICATOR_CAPS
'           Case 59138: PaneNum_Click        '* ID_INDICATOR_NUM
'           Case 59139: PaneScrl_Click       '* ID_INDICATOR_SCRL
         End Select
   End Select
End Sub

Private Sub StatusBar_PaneClick(ByVal Pane As XtremeCommandBars.StatusBarPane)
'   nStatusButton = nStatusButton
End Sub

Private Sub StatusBar_PaneDblClick(ByVal Pane As XtremeCommandBars.StatusBarPane)
   Select Case Pane.Id
'      Case 101:   PaneConexao_DblClick    '* CONEXÃO
      Case 102:   PaneUsusario_DblClick   '* USUARIO
'      Case 103:   PaneEquipe_DblClick     '* PERFIL/EQUIPE...
'      Case 104:   PaneStrech_DblClick     '* STRETCH
'      Case 108:   PaneData_DblClick       '* DATA
'      Case 109:   PaneHora_DblClick       '* HORA
      
'      Case 59137: PaneCapsLock_DblClick   '* ID_INDICATOR_CAPS
'      Case 59138: PaneNum_DblClick        '* ID_INDICATOR_NUM
'      Case 59139: PaneScrl_DblClick       '* ID_INDICATOR_SCRL
   End Select
   
End Sub
Private Sub PaneConexao_Click()
   Dim i       As Integer
   Dim Popup   As CommandBar
   Dim Control As CommandBarControl
   Dim nReturn As Integer
   Static nPar As Integer
     
   nPar = nPar + 1
   If nPar Mod 2 = 0 Then Exit Sub
     
   Set Popup = Me.CommandBars.Add("Popup", xtpBarPopup)
   
   While ReadIniFile(gLocalReg, "conection " & i, "ALIAS") <> ""
      Set Control = Popup.Controls.Add(xtpControlButton, i + 1, ReadIniFile(gLocalReg, "conection " & i, "ALIAS"))
      Control.Category = "POPUP_CONEXAO"
      Control.Parameter = "|CARREGADO=1"
      Control.Checked = (xVal(ReadIniFile(gLocalReg, "conections", "Last")) = i)
      i = i + 1
   Wend
   nReturn = Popup.ShowPopup(TPM_RETURNCMD) - 1
   
   MDI.MousePointer = vbHourglass
   Screen.MousePointer = vbHourglass
   If nReturn >= 0 Then
      If ReadIniFile(gLocalReg, "conections", "Last", 0) <> nReturn Then
         Call ReloadConnection(nReturn, False)
         Sys.IDLOJA = -1
         StatusBar.Pane(2).Text = ""
      End If
   End If
   Screen.MousePointer = vbDefault
   MDI.MousePointer = vbDefault
End Sub
Public Function ReloadConnection(IdConection As Integer, Optional bFull As Boolean = True) As Object
   Dim iConection As Integer
        
   iConection = ReadIniFile(gLocalReg, "conections", "Last", 0)
   Call WriteIniFile(gLocalReg, "conections", "Last", CStr(IdConection))
      
   gIDUSU = ""
   Sys.IDUSU = ""
   Set Sys.USER = Nothing
   If Not Splash Is Nothing Then Splash.IDUSU = ""
   
   Call MyLoadgCODSIS
   Call ExibeSenha(pTrocaConexao:=True)
   
   Call MDI.Reload(bFull)
   If Not (Sys.xdb.Conectado And Sys.xdb.Alias = ReadIniFile(gLocalReg, "Conection " & CStr(IdConection), "ALIAS")) Then
      Call WriteIniFile(gLocalReg, "conections", "Last", CStr(iConection))
   End If
   
   Set ReloadConnection = Sys
End Function
Private Sub PaneEquipe_Click()

End Sub
Private Sub PaneUsusario_DblClick() '* USUARIO
   Screen.MousePointer = vbHourglass
   
   Call Sys.xdb.Executa("Delete GLOGIN Where IDUSU = " & SqlStr(gIDUSU))
   If Not Splash Is Nothing Then Splash.IDUSU = ""
   gIDUSU = ""

   Sys.IDUSU = ""
   Set Sys.USER = Nothing
   Me.StatusBar.Pane(1).Text = ""
   Call ExibeSenha(True)
   Me.StatusBar.Pane(1).Text = gIDUSU
   
   'Call Sys.xDb.SrvDesconecta
   'gIDUSU = ""
   'Call ExibeSenha(True)
   If Not Splash Is Nothing Then
      If Not Splash.Cancelado Then
         Call MDI.Reload(False)
      End If
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub Timer_Timer()
   Dim nIndex As Integer
   Dim dData  As Date
   
   Static Segundo As Integer
   Static MeioSeg As Integer
   On Error Resume Next
   MeioSeg = IIf(Timer.Interval = 1000, 2, MeioSeg + 1)
   
   If MeioSeg >= 2 Then
      MeioSeg = 1
      Segundo = Segundo + 1
      If Segundo > 60 Then Segundo = 1
      
      If Segundo = 30 Or Segundo = 60 Then
         If Sys.xdb.Conectado Then
            dData = CDate(Sys.xdb.Sysdate(3))
         Else
            dData = Now()
         End If
         nIndex = Me.CommandBars.StatusBar.FindPane(108).Index
         If nIndex >= 0 Then
            Me.CommandBars.StatusBar(nIndex).Text = Format(dData, "dd/mm/yyyy")
         End If
         
         nIndex = Me.CommandBars.StatusBar.FindPane(109).Index
         If nIndex >= 0 Then
            Me.CommandBars.StatusBar(nIndex).Text = Format(dData, "hh:mm")
         End If
      End If
   End If
   
   If Timer.Interval = 500 Then
      If GetTag(Me.CommandBars.Tag, "EXIBIRRESULTADO", 0) = 1 Then
         Call ExibirMsgStatusBar
      End If
   End If
   If Sys.Propriedades("FCOMANDO") <> "" Then
      If Sys.Propriedades("FCOMANDO") = "End" Then
         End
      End If
   End If
   '* Consulta o valor da tecla por Api e se a resposta for -32767
   '* , então mostra tecla em Chr(i) e o valor de i
   'Dim i As Integer
   'For i = 0 To 255
   '   If GetAsyncKeyState(i) = -32767 Then
   '      If i = 123 Then MsgBox "Tecla F12"
   '   End If
   'Next
End Sub
Private Sub ExibirMsgStatusBar()
   Dim NumPisca As Integer
   Dim bMsgPositiva  As Boolean
   Dim oPane   As StatusBarPane
   Dim sMsg    As String
   Static nTime As Integer
   
   'On Error Resume Next
   
   NumPisca = GetTag(Me.CommandBars.Tag, "NUMPISCA", 0)
   bMsgPositiva = (GetTag(Me.CommandBars.Tag, "MSGPOSITIVA", 0) = 1)
   sMsg = GetTag(Me.CommandBars.Tag, "MSGRESULTADO", "")
   Set oPane = Me.CommandBars.StatusBar.FindPane(104)
      
   
   DoEvents
   Me.ProgressBar.Visible = False
   Me.ProgressBar.Value = 0
   oPane.Text = sMsg
   
   nTime = nTime + 1
   If NumPisca = 1 And nTime = 2 Then
      NumPisca = 0
   End If
   If NumPisca < 0 Then
   '   Me.PnlMessage.BackColor = IIf(MsgPositiva, COR.Verde, COR.Vermelho)
   '   Me.PnlMessage.ForeColor = COR.Branco
   '   Me.PnlMessage.Refresh
   '   Me.Refresh
   '   nTime = 0
   '   NumPisca = 0
   '   Unload Me
   End If
   If (nTime Mod 2) = 1 Then '* Cor
      oPane.TextColor = vbWhite
      oPane.BackgroundColor = IIf(bMsgPositiva, vbBlue, vbRed)
   Else '* Neutro
      oPane.TextColor = vbBlack
      oPane.BackgroundColor = Me.CommandBars.StatusBar.FindPane(109).BackgroundColor
   End If
   
   If nTime = NumPisca + 2 Then
      oPane.BackgroundColor = Me.CommandBars.StatusBar.FindPane(109).BackgroundColor
      oPane.TextColor = vbBlack
      oPane.Text = ""
      Me.ProgressBar.Visible = True
      Me.CommandBars.Tag = SetTag(Me.CommandBars.Tag, "EXIBIRRESULTADO", 0)
      Timer.Interval = 1000
      
      nTime = 0
      NumPisca = 0
   End If
End Sub
