VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MdiPrincipal 
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Compras Vs.1.0"
   ClientHeight    =   1155
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9285
   Icon            =   "MDITOOL.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDITOOL.frx":0442
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   900
      ButtonWidth     =   741
      ButtonHeight    =   741
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   21
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'01'"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'02'"
            Object.Tag             =   ""
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'03'"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'04'"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'05'"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'06'"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'07'"
            Object.Tag             =   ""
            Style           =   1
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'08'"
            Object.Tag             =   ""
            Style           =   2
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'09'"
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   1e-4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'10'"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'11'"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'12'"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'13'"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'14'"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'15'"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'16'"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'17'"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'18'"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'19'"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'20'"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "'21'"
            Object.ToolTipText     =   "Help"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BorderStyle     =   1
      OLEDropMode     =   1
      Begin VB.DriveListBox Drv1 
         Height          =   315
         Left            =   8760
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CmDialog 
      Left            =   2520
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer TmBD 
      Interval        =   60000
      Left            =   3000
      Top             =   480
   End
   Begin Crystal.CrystalReport RELAT 
      Left            =   3360
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   780
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15/05/01"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Data do Sistema"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "17:25"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Hora do Sistema"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Enabled         =   0   'False
            Text            =   "UUUUUUUUUU"
            TextSave        =   "UUUUUUUUUU"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Usuario do Sistema"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7329
            MinWidth        =   5292
            Text            =   "PROCESS"
            TextSave        =   "PROCESS"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Barra de Processamento"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   3069
            Text            =   "RIO DE JANEIRO - BR"
            TextSave        =   "RIO DE JANEIRO - BR"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Local do Sistema"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   34
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":0888
            Key             =   "TABLE"
            Object.Tag             =   "TABLE"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":0E6A
            Key             =   "TOOLS"
            Object.Tag             =   "TOOLS"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":144C
            Key             =   "PIECE"
            Object.Tag             =   "PIECE"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":15CE
            Key             =   "AREA"
            Object.Tag             =   "AREA"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":1750
            Key             =   "MERGE"
            Object.Tag             =   "MERGE"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":18D2
            Key             =   "CHAVE"
            Object.Tag             =   "CHAVE"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":1A54
            Key             =   "DELETE"
            Object.Tag             =   "DELETE"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":1BD6
            Key             =   "PROJET"
            Object.Tag             =   "PROJET"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":1D58
            Key             =   "TAB"
            Object.Tag             =   "TAB"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":1EDA
            Key             =   "DOLAR"
            Object.Tag             =   "DOLAR"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":205C
            Key             =   "EXIT"
            Object.Tag             =   "EXIT"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":23AA
            Key             =   "EDIT"
            Object.Tag             =   "EDIT"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":252C
            Key             =   "EXCEL"
            Object.Tag             =   "EXCEL"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":2B0E
            Key             =   "MAIL"
            Object.Tag             =   "MAIL"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":2C90
            Key             =   "EXPORT"
            Object.Tag             =   "EXPORT"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":2E12
            Key             =   "SEEK"
            Object.Tag             =   "SEEK"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":2F94
            Key             =   "FITA"
            Object.Tag             =   "FITA"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":3166
            Key             =   "GLOBO"
            Object.Tag             =   "GLOBO"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":32E8
            Key             =   "IMPORT"
            Object.Tag             =   "IMPORT"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":346A
            Key             =   "PRINT"
            Object.Tag             =   "PRINT"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":35EC
            Key             =   "MONEY"
            Object.Tag             =   "MONEY"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":376E
            Key             =   "OLHO"
            Object.Tag             =   "OLHO"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":3E7C
            Key             =   "SELECT"
            Object.Tag             =   "SELECT"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":3FFE
            Key             =   "PAPER"
            Object.Tag             =   "PAPER"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":4180
            Key             =   "REFRESH"
            Object.Tag             =   "REFRESH"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":4302
            Key             =   "TABLES"
            Object.Tag             =   "TABLES"
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":4484
            Key             =   "SAVE"
            Object.Tag             =   "SAVE"
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":4606
            Key             =   "SUPPLIER"
            Object.Tag             =   "SUPPLIER"
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":4BE8
            Key             =   "BACK"
            Object.Tag             =   "BACK"
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":4D6A
            Key             =   "HELP"
            Object.Tag             =   "HELP"
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":534C
            Key             =   "ADD"
            Object.Tag             =   "ADD"
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":54CE
            Key             =   "SUB"
            Object.Tag             =   "SUB"
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":5650
            Key             =   "BANCO"
            Object.Tag             =   "BANCO"
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":5C9A
            Key             =   "CLASSE"
            Object.Tag             =   "CLASSE"
         EndProperty
      EndProperty
   End
   Begin VB.Menu Mnu 
      Caption         =   "&Arquivo"
      Index           =   0
      Begin VB.Menu Mnu00 
         Caption         =   "&Abri Banco de Dados"
         Index           =   0
         Begin VB.Menu Mnu0000 
            Caption         =   "&Microsoft Access.."
            Index           =   0
         End
         Begin VB.Menu Mnu0000 
            Caption         =   "&ODBC..."
            Index           =   1
         End
      End
      Begin VB.Menu Mnu00 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu Mnu00 
         Caption         =   "&Imprimir..."
         Index           =   2
      End
      Begin VB.Menu Mnu00 
         Caption         =   "&Configurar Impressão..."
         Index           =   3
      End
      Begin VB.Menu Mnu00 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu Mnu00 
         Caption         =   "&Sair"
         Index           =   5
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "&Feramentas"
      Index           =   1
      Begin VB.Menu Mnu01 
         Caption         =   "&Criar Arquivo .Cls"
         Index           =   0
      End
      Begin VB.Menu Mnu01 
         Caption         =   "&Analisar Projeto"
         Index           =   1
      End
      Begin VB.Menu Mnu01 
         Caption         =   "&Montar Formulário"
         Index           =   2
      End
      Begin VB.Menu Mnu01 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu Mnu01 
         Caption         =   "C&onfigurar"
         Index           =   4
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "&Segurança..."
      Index           =   2
      Begin VB.Menu Mnu02 
         Caption         =   "&Grupo/Usuário/Permissão"
         Index           =   0
      End
      Begin VB.Menu Mnu02 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu Mnu02 
         Caption         =   "&Troca Usuário"
         Index           =   2
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "&Janela"
      Index           =   3
      WindowList      =   -1  'True
      Begin VB.Menu Mnu03 
         Caption         =   "&Cascata"
         Index           =   0
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "&Ajuda"
      Index           =   4
      Begin VB.Menu Mnu04 
         Caption         =   "&Contexto"
         Index           =   0
      End
      Begin VB.Menu Mnu04 
         Caption         =   "&Índice"
         Index           =   1
      End
      Begin VB.Menu Mnu04 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu Mnu04 
         Caption         =   "&Sobre o Sistema..."
         Index           =   3
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MnuMouse 
      Caption         =   "Mouse TreeProj"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu MnuMouse01 
         Caption         =   "Propriedades"
         Index           =   0
      End
   End
End
Attribute VB_Name = "MdiPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PrimeiraVez%
Dim f As Form
Public Sub F_EXCLUIR()
   If Not MDIFilho Is Nothing Then
      On Error Resume Next
      MDIFilho.F_EXCLUIR
      Select Case Err
         Case 438:  On Error GoTo 0
      End Select
    End If
End Sub
Public Sub F_INCLUIR()
   If Not MDIFilho Is Nothing Then
      On Error Resume Next
      MDIFilho.F_INCLUIR
      Select Case Err
         Case 438:  On Error GoTo 0
      End Select
    End If
End Sub

Public Sub F_OPCAO(Op$)
'   Dim ClsSeg As New DS_SEGURANCA
   Dim Frm As Form
   Dim Sql$, Arq$
   
   Select Case Op$
'******Case "00": 'CADASTRO
         Case "0000" 'Abrir Banco de Dados
            Case "000000"
               With DB
                  FrmOpBanco.Show vbModal
                  If Not Sys.isODBC Then
                     If Dir("C:\DSR\", vbDirectory) <> "" Then
                        SysMdi.CmDialog.InitDir = "C:\DSR\"
                     End If
                     Arq$ = .Alias
                     If .Alias = "" Then
                        Arq$ = ProcurarArquivo(SysMdi.CmDialog, "Abrir Banco de Dados Access", , "Microsoft Access MDBs (*.mdb)|*.mdb")
                        .isODBC = False
                        .dbDrive = SysMdi.CmDialog.Tag
                        .dbName = Arq$
                     End If
                     If Arq$ <> "" Then
                        Call .SrvConecta(.dbDrive, .dbName, "", "", "", "")
                     End If
                  Else
                     .isODBC = True
                     frmODBCLog.Show vbModal
                  End If
               End With
               
               For i = 2 To 5
                  cAux = Trim(GetSetting(Sys.AppName, "Outros", "BDRecente" & CStr(i - 1), ""))
                  If DB.Alias <> cAux And cAux <> "" Then
                     Call SaveSetting(Sys.AppName, "Outros", "BDRecente" & CStr(i), cAux)
                  End If
               Next
               Call SaveSetting(Sys.AppName, "Outros", "BDRecente1", DB.Alias)
            Case "000001"
               frmODBCLog.Show vbModal
'               DB.isODBC = "S"
'               Call DB.SrvConecta(dbdrive, dbName, "", "", "", "")
         Case "0001": 'Separador
         Case "0002": FrmPrint.Show vbModeless
         Case "0003": Call PrintSetup(SysMdi.CmDialog)
         Case "0004": 'Separador
         Case "0005": Unload Me
      
'******Case "01":FERRAMENTAS
         Case "0100"
'            If Not DB.Conectado Then Call F_OPCAO("000000")
'            Screen.MousePointer = vbHourglass
'            If DB.Conectado Then
               FrmMontaCls.Show vbModal
               Set FrmMontaCls = Nothing
'            End If
         Case "0101": FrmDescProj.Show vbModeless
         Case "0102"
'            If Not DB.Conectado Then Call F_OPCAO("000000")
'            If DB.Conectado Then
               FrmWizPrj.Show vbModeless
'            End If
         Case "0103": '* Separator
         Case "0104": '* FrmConfig.Show vbModal
'*****Case "02":SEGURANÇA
         Case "0200"
         Case "0201": 'Separador
         Case "0202"
            If UCase(ClsSeg.IDUSER) = "FIM" Then Exit Sub
            If SysMdi.Toolbar.AllowCustomize Then
              SysMdi.Toolbar.RestoreToolbar AppTitle_Sistema, ANALISTA, SysMdi.Toolbar.Name
            End If
            With SysMdi.StatusBar
              .Panels(3).Text = IIf(Len(Trim(ANALISTA)) = 0, "DSVM", ANALISTA)
              .Panels(4).Text = SUP_Nome_Filial$
              .Panels(5).Text = DPH_Territorio$
            End With
            SysMdi.StatusBar.Refresh
            Call SetDefault(hWnd)

'******Case "03":JANELA
         Case "0300": Me.Arrange 0
      
'******Case "04":AJUDA
         Case "0400": 'Call AboutShow 'frmSobre.Show
         Case "0401": 'Call AboutShow 'frmSobre.Show
         Case "0402": 'Separator
         Case "0403": Call AboutShow(App) 'frmSobre.Show
     End Select
     Screen.MousePointer = vbDefault
End Sub

Public Sub F_PROCURAR()
   If Not MDIFilho Is Nothing Then
      On Error Resume Next
      MDIFilho.F_PROCURAR
      Select Case Err
         Case 438:  On Error GoTo 0
      End Select
    End If
End Sub

Public Sub F_REFRESH()
   If Not MDIFilho Is Nothing Then
      On Error Resume Next
      MDIFilho.F_REFRESH
      Select Case Err
         Case 438:  On Error GoTo 0
      End Select
   End If
End Sub

Public Sub F_SALVAR()
   If Not MDIFilho Is Nothing Then
      On Error Resume Next
      MDIFilho.F_SALVAR
      Select Case Err
         Case 438:  On Error GoTo 0
      End Select
    End If
End Sub

Public Function F_VOLTAR()
   On Error Resume Next
   If FORMS.Count = 1 Then
      Unload Me
   Else
      If MDIFilho.Name = "FrmCad" Then
         MDIFilho.Visible = False
      Else
         Unload MDIFilho
      End If
   End If
End Function
Private Sub MDIForm_Activate()
   Call SetHourglass(hWnd)
'   Me.ZOrder 1
   If Flag_Inicio% Then
    
      With SysMdi
         .Caption = App.Title + " " + LoadMsg(49) + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor))
         .WindowState = vbMaximized
         .TmBD.Enabled = True
         .Visible = True
      End With
   '====voltar
      Call ToolbarPrincipal
      If SysMdi.Toolbar.AllowCustomize Then
'         SysMdi.Toolbar.RestoreToolbar AppTitle_Sistema, ANALISTA, SysMdi.Toolbar.Name
      End If
         DPH_Territorio$ = "RJ - RIO DE JANEIRO "
         Sys_CodLocal$ = "RJ"
         Sys_DscLocal$ = "RIO DE JANEIRO"
         Sys_IdPais$ = "BR"
         Sys_DscPais$ = "BRASIL"
      Call DB.SrvDesconecta
      With SysMdi.StatusBar
         .Panels(3).Text = IIf(Len(Trim(ANALISTA)) = 0, "DSVM", ANALISTA)
         .Panels(4).Text = SUP_Nome_Filial$
         .Panels(5).Text = DPH_Territorio$
      End With
      Flag_Inicio% = False
      flag_inicio_senha% = False
   End If
   Call SetDefault(hWnd)
End Sub
Private Sub MDIForm_Initialize()
   Call AutoInstalacao(App)
End Sub

Private Sub MDIForm_Load()
   Dim ClsLoad As New DS_LOAD
   Call SetHourglass(hWnd)
   With ClsLoad
      .AnoDSVM = "Maio - 1999"
      .Aplic = App
      .Show
   End With
   Call Load
   Call Active
   Set ClsLoad = Nothing
   Me.Visible = True
   Call SetDefault(hWnd)
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'   If vbNo= ExibirPergunta(LoadMsg(30), Sys.AppName + " " + Mid(Sys.AppVer, InStr(Sys.AppVer, ".") - 1)) Then
'      Cancel = True
'   Else
'      Cancel = False
'  End If
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
   Call SetHourglass(hWnd)
   If SysMdi.Toolbar.AllowCustomize Then
      SysMdi.Toolbar.SaveToolbar AppTitle_Sistema, ANALISTA, SysMdi.Toolbar.Name
   End If
   For Each Form In FORMS
      If Form Is Form Then
         Unload Form
         Set Form = Nothing
         Exit For
       End If
   Next Form
   Call DB.SrvDesconecta
   End
End Sub
Private Sub Mnu_Click(Index As Integer)
   Call F_OPCAO(StrZero(Index, 2))
End Sub
Private Sub Mnu00_Click(Index As Integer)
   Call F_OPCAO("00" + StrZero(Index, 2))
End Sub
Private Sub Mnu0000_Click(Index As Integer)
   Call F_OPCAO("0000" + StrZero(Index, 2))
End Sub
Private Sub Mnu01_Click(Index As Integer)
   Call F_OPCAO("01" + StrZero(Index, 2))
End Sub
Private Sub Mnu02_Click(Index As Integer)
   Call F_OPCAO("02" + StrZero(Index, 2))
End Sub
Private Sub Mnu03_Click(Index As Integer)
   Call F_OPCAO("03" + StrZero(Index, 2))
End Sub
Private Sub Mnu04_Click(Index As Integer)
   Call F_OPCAO("04" + StrZero(Index, 2))
End Sub
Private Sub TmBD_Timer()
    DB.MinutosConexao = DB.MinutosConexao + 1
    If DB.MinutosConexao = (TEMPO_MAXIMO_CONEXAO% + 1) Then
       Call DB.DerrubarSistema
       End
    End If
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
   Dim Frm As Form, Ind%
   Ind% = Val(Mid(Button.Key, 2, 2))
   Select Case Ind%
      Case BT_BANCO: Call F_OPCAO("000000")
      Case BT_CLS: Call F_OPCAO("0100")
      Case BT_PRJ: Call F_OPCAO("0101")
      
      Case BT_TABELAS: Call F_OPCAO("0102")
      Case 5:
      Case 6:
      Case 7:
      '==============================
      '==============================
      Case BT_RPT: Call F_OPCAO("0005")
      Case 9:
      Case BT_REFRESH: Call F_OPCAO("0100")
      Case BT_FIND: Call F_OPCAO("0101")
      Case BT_SAVE: Call F_OPCAO("0102")
      Case BT_DEL: Call F_OPCAO("0103")
      Case BT_VOLTAR: Call F_VOLTAR
      Case BT_SAIR: Unload Me
      Case BT_PAIS: FrmPais.Show
      Case BT_USU: Call F_OPCAO("0400")
      Case BT_UTILIT: 'Set Frm = FrmTools
      Case BT_BACKUP: 'FrmBackup.Show vbModal
      Case BT_ABOUT: Call F_OPCAO("0603")
   End Select
End Sub

Public Sub Load()
    Dim Txt$, AnoDSVM$
    Dim estilo%, ERRO%
    Set SysMdi = MdiPrincipal
    Call SetHourglass(hWnd)

'=============================
'=============================
    Sys.Appexe = App.EXEName  '"SCC"
    Sys.AppName = App.Title    '"SisCompra"
    Sys.AppTitle = App.ProductName   '"Sistema de Compras"
    Sys.AppVer = LoadRes("S49") + Trim(CStr(App.Major)) + "." + Trim(CStr(App.Minor))
    AppDate$ = Format$(Now, "dd/mm/yyy - hh:mm:ss")
    SYS_Nome_Empresa$ = App.LegalCopyright  '"Delphos Serviços Técnicos S/A"
    
'=============================
'=============================
   Call MakePath("C:\TMP\")
'=============================
'=============================
   Call GetConfig           'Variáveis do Registro do Sistema
   'Cria Banco de Dados se ele não existir e se o banco em uso for ACCESS
   If InArray(Val(dbVersion), Array(1, 8, 16, 32)) Then
      If Not FileExists(dbDrive & dbName) Then Call CriarBD(dbDrive & dbName)
'      criartabela "teste", dbdrive & dbName
   End If
    'testa se já existe uma cópia da aplicação rodando
    If App.PrevInstance Then
        SysMdi.Caption = Sys.AppName$ + " [Cópia]" 'altera caption da cópia
        On Error Resume Next
        'tenta ativar aplicação que já estava rodando (caption de MDIForm)
        AppActivate Sys.Appexe$
        ERRO% = Err
        On Error GoTo 0
        'testa erro (aplicação pode estar com o foco em um form não MDIChild)
        If ERRO% <> 0 Then
           Txt$ = "Já existe uma cópia da aplicação rodando."
           Txt$ = Txt$ + Chr$(10) + "Pressione Ok para fechar esta mensagem"
           Txt$ = Txt$ + Chr$(10) + "e Alt+Tab para localizar a aplicação."
            estilo% = vbSystemModal + vbExclamation
            Screen.MousePointer = vbDefault
            MsgBox Txt$, estilo%, SysMdi.Caption
            DoEvents
        Else
            'maximiza aplicação que já estava rodando
            SendKeys "% X"
        End If
        'encerra aplicação cópia
        End
    End If
    flag_inicio_senha% = True
    Flag_Inicio% = True
    MsgTitulo$ = Sys.AppName$
End Sub
Public Sub Active()
   Dim Sql$
'   Dim ClsSeg As New SEGURANCA
   Call SetHourglass(hWnd)
   If flag_inicio_senha% And Flag_Inicio% Then
      'Configura arquivo de help
      'App.HelpFile = DIRETORIO_REDE$ + "materiais.hlp"
        
      flag_inicio_senha% = False
      
'      Call DPH_Init
      '===================
      '===================
      Call DB.SrvConecta(dbDrive, dbName, "", "", "", "")
      ClsSeg.dBase = DB
      ClsSeg.CodSys = Sys.Appexe
      MICRO = "DSVM"
      DoEvents
      'Salva mês, dia e hora de entrada no sistema
      Call SetDefault(hWnd)
    End If
End Sub

