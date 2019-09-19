VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmConfig 
   AutoRedraw      =   -1  'True
   Caption         =   "Configuração"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtMicro 
      Height          =   330
      Left            =   2280
      LinkTimeout     =   30
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      WhatsThisHelpID =   10541
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6165
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Contrutor"
      TabPicture(0)   =   "FrmConfig.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LstSetup"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Geral"
      TabPicture(1)   =   "FrmConfig.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Lbl(7)"
      Tab(1).Control(1)=   "Lbl(0)"
      Tab(1).Control(2)=   "Lbl(1)"
      Tab(1).Control(3)=   "CmdDrv(2)"
      Tab(1).Control(4)=   "TxtDrvErro"
      Tab(1).Control(5)=   "Frame1"
      Tab(1).Control(6)=   "ScrFundo"
      Tab(1).Control(7)=   "CmbIdioma"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Modelos"
      TabPicture(2)   =   "FrmConfig.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.ListBox LstSetup 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         ItemData        =   "FrmConfig.frx":0054
         Left            =   240
         List            =   "FrmConfig.frx":0061
         Style           =   1  'Checkbox
         TabIndex        =   13
         Top             =   600
         Width           =   3015
      End
      Begin VB.ComboBox CmbIdioma 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmConfig.frx":00B4
         Left            =   -73320
         List            =   "FrmConfig.frx":00C4
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         WhatsThisHelpID =   10522
         Width           =   1305
      End
      Begin VB.VScrollBar ScrFundo 
         Height          =   2580
         Left            =   -73680
         TabIndex        =   8
         Top             =   675
         WhatsThisHelpID =   10523
         Width           =   255
      End
      Begin VB.Frame Frame1 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   7
         Top             =   600
         WhatsThisHelpID =   10524
         Width           =   1215
         Begin VB.Image ImgFundos 
            Height          =   690
            Index           =   3
            Left            =   600
            Picture         =   "FrmConfig.frx":00EE
            Top             =   1560
            WhatsThisHelpID =   10516
            Width           =   690
         End
         Begin VB.Image ImgFundos 
            Height          =   690
            Index           =   2
            Left            =   240
            Picture         =   "FrmConfig.frx":0DD0
            Top             =   1920
            WhatsThisHelpID =   10517
            Width           =   690
         End
         Begin VB.Shape ShpFundo 
            BorderWidth     =   2
            Height          =   855
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   180
            Width           =   975
         End
         Begin VB.Image ImgFundos 
            Height          =   690
            Index           =   1
            Left            =   240
            Picture         =   "FrmConfig.frx":273A
            Top             =   1080
            WhatsThisHelpID =   10518
            Width           =   690
         End
         Begin VB.Image ImgFundos 
            Height          =   690
            Index           =   0
            Left            =   240
            Picture         =   "FrmConfig.frx":341C
            Top             =   240
            WhatsThisHelpID =   10519
            Width           =   690
         End
      End
      Begin VB.TextBox TxtDrvErro 
         Enabled         =   0   'False
         Height          =   330
         Left            =   -73320
         LinkTimeout     =   30
         TabIndex        =   6
         Top             =   1200
         WhatsThisHelpID =   10528
         Width           =   3555
      End
      Begin VB.CommandButton CmdDrv 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   " ..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -69720
         TabIndex        =   5
         Top             =   1200
         WhatsThisHelpID =   10527
         Width           =   315
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Idioma"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   -73320
         TabIndex        =   12
         Top             =   360
         WhatsThisHelpID =   10539
         Width           =   540
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Fundo de Tela"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   -74880
         TabIndex        =   11
         Top             =   360
         WhatsThisHelpID =   10540
         Width           =   1125
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Drive Erro : "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   7
         Left            =   -73320
         TabIndex        =   10
         Top             =   960
         WhatsThisHelpID =   10530
         Width           =   1020
      End
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   405
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   3840
      WhatsThisHelpID =   10543
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   714
      _StockProps     =   78
      ForeColor       =   255
      Picture         =   "FrmConfig.frx":40FE
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   405
      Index           =   0
      Left            =   3960
      TabIndex        =   3
      Top             =   3840
      WhatsThisHelpID =   10542
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   714
      _StockProps     =   78
      Picture         =   "FrmConfig.frx":4B38
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Micro"
      Height          =   195
      Index           =   6
      Left            =   2640
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      WhatsThisHelpID =   10544
      Width           =   390
   End
End
Attribute VB_Name = "FrmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ArrScr
Public CarregandoConfig
Public Function ValidaCampos()
   ValidaCampos = False
   '   If Right(Me.TxtDbDrive, 1) <> "\" Then Me.TxtDbDrive = Me.TxtDbDrive & "\"
   '   If Right(Me.TxtDrvRpt, 1) <> "\" Then Me.TxtDrvRpt = Me.TxtDrvRpt & "\"
   '   If Trim$(Me.TxtDbDrive) = "" Or Not FileExists(Me.TxtDbDrive + Me.TxtDbName) Then
   '      Call ExibirAviso(LoadRes("S27"), LoadRes("S1"))
   '      Call Set_Focus(Me.TxtDbDrive)
   '      Exit Function
   '   End If
   '   If Trim$(Me.TxtDrvRpt) = "" Then
   '      Call ExibirAviso(LoadRes("S27"), LoadRes("S1"))
   '      Call Set_Focus(Me.TxtDrvRpt)
   '      Exit Function
   '   End If
   '   If Trim$(Me.TxtDbName) = "" Or Not FileExists(Me.TxtDbDrive + Me.TxtDbName) Then
   '      Call ExibirAviso(LoadRes("S27"), LoadRes("S1"))
   '      Call Set_Focus(Me.TxtDbName)
   '      Exit Function
   '   End If
   ValidaCampos = True
End Function
'Public Sub SaveConfig()
'   Dim lODBC$, lVersao$, lName$

'   lODBC = Me.OptODBC(0)
'   lVersao = UCase(Me.CmbVersao.List(Me.CmbVersao.ListIndex))
'   lName = UCase(Me.TxtDbName)

'*** [ Database Format ] ***
'   Call SaveSetting(Sys.AppName, "Database Format", "DBODBC", lODBC$)
'   Call SaveSetting(Sys.AppName, "Database Format", "DBVERSAO", lVersao$)
'   Call SaveSetting(Sys.AppName, "Database Format", "DBNAME", lName$)

'*** [ Database Drive ] ***
'   Call SaveSetting(Sys.AppName, "Database Drive", "DBDRIVE", Me.TxtDbDrive)
'   Call SaveSetting(Sys.AppName, "Database Drive", "DRVRPT", Me.TxtDrvRpt)

'*** [ Setup ] ***
'   Call SaveSetting(Sys.AppName, "Setup", "MICRO", Me.TxtMicro)
'   Call SaveSetting(Sys.AppName, "Setup", "FUNDOTELA", Me.ShpFundo.Tag)
'   Call SaveSetting(Sys.AppName, "Setup", "IDIOMA", Me.CmbIdioma)
'   Call SaveSetting(Sys.AppName, "Setup", "DrvErro", Me.TxtDrvErro)
'
'End Sub
Private Sub CmdDrv_Click(Index As Integer)
   Select Case Index
      '      Case 0: Call GetPath(hwnd, Me.TxtDbDrive)
      '      Case 1: Call GetPath(hwnd, Me.TxtDrvRpt)
      Case 2: Call GetPath(hwnd, Me.TxtDrvErro)
   End Select
End Sub
Private Sub CmdOper_Click(Index As Integer)
   Select Case Index
      Case 0
         If ValidaCampos() Then
            Call UpdateConfig
            Call SaveConfig
         Else
            Exit Sub
         End If
         '      Case 1: BANCO.TB_CONFIG.Cancelado = True
   End Select
   Unload Me
End Sub

Private Sub Form_Activate()
   Call CarregaConfig
   '   Set sys.MDIFilho = FrmConfig
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   KeyAscii = SendTab(Me, KeyAscii)
End Sub

Private Sub Form_Load()
   Dim i%, Pos%
   Me.Caption = "Configurar " & App.ProductName
   Call ConfigForm(Me, SysMdi.Icon, Sys.Proj.FundoTela)

   '*** [ Setup ] ***
   Me.LstSetup.Selected(0) = GetSetting(Sys.Proj.AppName, "Setup", "LoadIni", False)
   Me.LstSetup.Selected(1) = GetSetting(Sys.Proj.AppName, "Setup", "ExibeSubPasta", False)
   Me.LstSetup.Selected(2) = GetSetting(Sys.Proj.AppName, "Setup", "SalvarOnLine", False)

   Me.TxtMicro = Sys.Proj.MICRO
   Select Case GetSetting(Sys.Proj.AppName, "Setup", "IDIOMA", "Português")
      Case "Português": Me.CmbIdioma.ListIndex = 0
      Case "Inglês": Me.CmbIdioma.ListIndex = 1
      Case "Francês": Me.CmbIdioma.ListIndex = 2
      Case "Espanhol": Me.CmbIdioma.ListIndex = 2
   End Select

   '   If ClsUser.Grp = GRPANALISTA Then
   '      Me.OptODBC(0).Enabled = True
   '      Me.OptODBC(1).Enabled = True
   '      Me.CmbVersao.Enabled = True
   '      Me.Lbl(6).Visible = True
   '      Me.TxtMicro.Visible = True
   '      Me.TxtDbName.Enabled = True
   '      Me.TxtDbDrive.Enabled = True
   '      Me.TxtDrvRpt.Enabled = True
   '
   '      Me.CmbIdioma.Enabled = True
   '   End If
   '*** [ Database Format ] ***
   '   Me.OptODBC(IIf(Sys.dbODBC, 0, 1)).Value = True
   '   Select Case UCase(Sys.dbVersao)
   '      Case "ACCESS1.0": i% = 0
   '      Case "ACCESS1.1": i% = 1
   '      Case "ACCESS2.0": i% = 2
   '      Case "ACCESS3.0": i% = 3
   '      Case Else: i% = 0
   '   End Select
   '   Me.CmbVersao.ListIndex = i%

   '   Me.TxtDbName = Sys.dbName

   '*** [ Database Drive ] ***
   '   Me.TxtDbDrive = Sys.dbDrive
   '   Me.TxtDrvRpt = Sys.DrvRpt

   '*****************************
   i = 4
   ReDim ArrScr(i%)
   Pos = ImgFundos(0).Height + 120
   For i = 0 To UBound(ArrScr)
      ArrScr(i) = Pos + (i * Pos)
   Next

   Me.ScrFundo.Min = Me.ImgFundos.LBound
   Me.ScrFundo.Max = Me.ImgFundos.UBound

   If Not IsEmpty(ArrScr) Then
      ScrFundo.Value = 1
      ScrFundo.Value = 0
   End If

   Me.ScrFundo.Enabled = (Me.ImgFundos.UBound >= 2)
   Me.ScrFundo.Value = Val(Right(Sys.Proj.FundoTela, 1))
   Me.TxtDrvErro = Sys.Proj.DrvErro
   '*****************************
   'sys.dbDrive_Orig = sys.dbDrive

   Call ConfigForm(Me, SysMdi.Icon, Sys.Proj.FundoTela)
   Call SetDefault(hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '  Set sys.MDIFilho = Nothing
End Sub

Private Sub ImgFundos_Click(Index As Integer)
   Me.ScrFundo.Value = Index
End Sub

Private Sub ScrFundo_Change()
   Dim Opt As Object, Scr As Control
   Dim x%, i%, Bool%

   Set Opt = ImgFundos
   Set Scr = ScrFundo

   If IsEmpty(ArrScr) Then Exit Sub
   If UBound(ArrScr) = 0 Then Exit Sub
   x% = Scr.Value
   For i = 0 To x%
      Opt(IIf(i - 1 < 0, 0, i - 1)).Visible = False
   Next
   For i = Opt.LBound To Opt.UBound
      If i% + x% > Opt.UBound Then Exit For
      Bool = IIf(i% + x% < 0 Or i% + x% > Opt.UBound + 1, False, True)
      Opt(i% + x%).Visible = Bool
      Opt(i% + x%).Move Me.ImgFundos(0).Left, ArrScr(i%) - Me.ImgFundos(0).Height + 120
   Next

   For i = Opt.LBound To Opt.UBound
      If Opt(i).Top = Opt(0).Top Then
         If Opt(i).Visible Then Exit For
      End If
   Next
   ShpFundo.Tag = "FUNDO" & IIf(i = 0, "", CStr(i))
   Call PintarFundo(Me, ShpFundo.Tag)
   Me.Refresh
End Sub
Private Sub TxtDbDrive_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtDbName_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtDrvRpt_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtMicro_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub UpdateConfig()
   Dim i%, n As Variant
   With Sys
      With .Constru
         .LoadIni = Me.LstSetup.Selected(0)
         .ExibeSubPasta = Me.LstSetup.Selected(1)
         .SalvarOnLine = Me.LstSetup.Selected(2)
      End With
   End With
End Sub
Public Sub CarregaConfig()
   CarregandoConfig = True
   With Sys
      With .Constru
         Me.LstSetup.Selected(0) = .LoadIni
         Me.LstSetup.Selected(1) = .ExibeSubPasta
         Me.LstSetup.Selected(2) = .SalvarOnLine
      End With
   End With
End Sub
