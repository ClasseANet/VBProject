VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.MDIForm MdiPrincipal 
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Utilitários Tecaplus Vs.1.0"
   ClientHeight    =   1155
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9285
   Icon            =   "MDITOOL.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDITOOL.frx":0442
   StartUpPosition =   3  'Windows Default
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
            TextSave        =   "22/12/1999"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Data do Sistema"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "20:35"
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
         NumListImages   =   35
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":0888
            Key             =   "TABLE"
            Object.Tag             =   "TABLE"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":0E6A
            Key             =   "SQL"
            Object.Tag             =   "SQL"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":0FEE
            Key             =   "TOOLS"
            Object.Tag             =   "TOOLS"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":15D0
            Key             =   "PIECE"
            Object.Tag             =   "PIECE"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":1752
            Key             =   "AREA"
            Object.Tag             =   "AREA"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":18D4
            Key             =   "MERGE"
            Object.Tag             =   "MERGE"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":1A56
            Key             =   "CHAVE"
            Object.Tag             =   "CHAVE"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":1BD8
            Key             =   "DELETE"
            Object.Tag             =   "DELETE"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":1D5A
            Key             =   "PROJET"
            Object.Tag             =   "PROJET"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":1EDC
            Key             =   "TAB"
            Object.Tag             =   "TAB"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":205E
            Key             =   "DOLAR"
            Object.Tag             =   "DOLAR"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":21E0
            Key             =   "EXIT"
            Object.Tag             =   "EXIT"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":252E
            Key             =   "EDIT"
            Object.Tag             =   "EDIT"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":26B0
            Key             =   "EXCEL"
            Object.Tag             =   "EXCEL"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":2C92
            Key             =   "MAIL"
            Object.Tag             =   "MAIL"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":2E14
            Key             =   "EXPORT"
            Object.Tag             =   "EXPORT"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":2F96
            Key             =   "SEEK"
            Object.Tag             =   "SEEK"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":3118
            Key             =   "FITA"
            Object.Tag             =   "FITA"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":32EA
            Key             =   "GLOBO"
            Object.Tag             =   "GLOBO"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":346C
            Key             =   "IMPORT"
            Object.Tag             =   "IMPORT"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":35EE
            Key             =   "PRINT"
            Object.Tag             =   "PRINT"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":3770
            Key             =   "MONEY"
            Object.Tag             =   "MONEY"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":38F2
            Key             =   "OLHO"
            Object.Tag             =   "OLHO"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":4000
            Key             =   "SELECT"
            Object.Tag             =   "SELECT"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":4182
            Key             =   "PAPER"
            Object.Tag             =   "PAPER"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":4304
            Key             =   "REFRESH"
            Object.Tag             =   "REFRESH"
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":4486
            Key             =   "TABLES"
            Object.Tag             =   "TABLES"
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":4608
            Key             =   "SAVE"
            Object.Tag             =   "SAVE"
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":478A
            Key             =   "SUPPLIER"
            Object.Tag             =   "SUPPLIER"
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":4D6C
            Key             =   "BACK"
            Object.Tag             =   "BACK"
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":4EEE
            Key             =   "HELP"
            Object.Tag             =   "HELP"
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":54D0
            Key             =   "ADD"
            Object.Tag             =   "ADD"
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":5652
            Key             =   "SUB"
            Object.Tag             =   "SUB"
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":57D4
            Key             =   "BANCO"
            Object.Tag             =   "BANCO"
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDITOOL.frx":5E1E
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
            Checked         =   -1  'True
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
      Begin VB.Menu Mnu00 
         Caption         =   "-"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu00 
         Caption         =   ""
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu00 
         Caption         =   ""
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu00 
         Caption         =   ""
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu00 
         Caption         =   ""
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu00 
         Caption         =   ""
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu00 
         Caption         =   ""
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu00 
         Caption         =   ""
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu00 
         Caption         =   ""
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu00 
         Caption         =   ""
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu00 
         Caption         =   ""
         Index           =   16
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "&Editar"
      Index           =   1
      Begin VB.Menu Mnu01 
         Caption         =   "&Refresh"
         Index           =   0
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "&Ferramentas"
      Index           =   2
      Begin VB.Menu Mnu02 
         Caption         =   "&Localizar/Analizar Carga"
         Index           =   0
      End
      Begin VB.Menu Mnu02 
         Caption         =   "&Executar ""Query"""
         Index           =   1
      End
      Begin VB.Menu Mnu02 
         Caption         =   "C&onfigurar"
         Index           =   2
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "&Segurança..."
      Index           =   3
      Begin VB.Menu Mnu03 
         Caption         =   "&Grupo/Usuário/Permissão"
         Index           =   0
      End
      Begin VB.Menu Mnu03 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu Mnu03 
         Caption         =   "&Troca Usuário"
         Index           =   2
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "&Janela"
      Index           =   4
      WindowList      =   -1  'True
      Begin VB.Menu Mnu04 
         Caption         =   "&Cascata"
         Index           =   0
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "&Ajuda"
      Index           =   5
      Begin VB.Menu Mnu05 
         Caption         =   "&Contexto"
         Index           =   0
      End
      Begin VB.Menu Mnu05 
         Caption         =   "&Índice"
         Index           =   1
      End
      Begin VB.Menu Mnu05 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu Mnu05 
         Caption         =   "&Sobre o Sistema..."
         Index           =   3
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MnuMousePrincipal 
      Caption         =   "Mouse "
      Begin VB.Menu MnuMouse 
         Caption         =   "LstTabela [ MnuMouse(0) ]"
         Index           =   0
         Begin VB.Menu MnuMouse00 
            Caption         =   "Abrir Tabela"
            Index           =   0
         End
         Begin VB.Menu MnuMouse00 
            Caption         =   "Exibir Tabelas de Sistema"
            Index           =   1
         End
      End
      Begin VB.Menu MnuMouse 
         Caption         =   "GrdView [ MnuMouse(1) ]"
         Index           =   1
         Begin VB.Menu MnuMouse01 
            Caption         =   "&Imprimir"
            Index           =   0
         End
         Begin VB.Menu MnuMouse01 
            Caption         =   "&Ordenação Múltipla"
            Index           =   1
         End
      End
   End
End
Attribute VB_Name = "MdiPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PrimeiraVez%
Dim f As Form
Public Sub F_EXCLUIR()
   If Not Sys.MDIFilho Is Nothing Then
      On Error Resume Next
      Sys.MDIFilho.F_EXCLUIR
      Select Case Err
         Case 438:  On Error GoTo 0
      End Select
    End If
End Sub
Public Sub F_INCLUIR()
   If Not Sys.MDIFilho Is Nothing Then
      On Error Resume Next
      Sys.MDIFilho.F_INCLUIR
      Select Case Err
         Case 438:  On Error GoTo 0
      End Select
    End If
End Sub

Public Sub F_OPCAO(Op$)
   Dim ClsSeg As New SEGURANCA
   Dim Frm As Form
   Dim Sql$, Arq$, i%
   
   Select Case Op$
'******Case "00": 'CADASTRO
         Case "0000" 'Abrir Banco de Dados
            Case "000000"
               With DB
                  dbODBC = "S"
                  .dbODBC = "S"
                  frmODBCLog.Show vbModal
                  BANCO.dBase = DB
                  Me.Mnu00(6).Visible = True
                  For i = 1 To DB.dBases.Count
                     Me.Mnu00(6 + i).Caption = DB.dBases(i).Name
                     Me.Mnu00(6 + i).Visible = True
                     Me.Mnu00(6 + i).Checked = (DB.dBases(i).Name = BANCO.dBase.dBase.Name)
                  Next
               End With
               Call F_REFRESH
            Case "000001": frmODBCLog.Show vbModal
         Case "0001": 'Separador
         Case "0002": 'FrmPrint.Show vbModeless
         Case "0003": Call PrintSetup(SysMdi.CmDialog)
         Case "0004": 'Separador
         Case "0005": Unload Me

'******Case "01": EDIT
         Case "0100": Call F_REFRESH
      
'******Case "02":FERRAMENTAS
         Case "0200"
            If Not DB.Conectado Then Call F_OPCAO("000000")
            If DB.Conectado Then FrmSeekCrg.Show vbModeless
         Case "0201": FrmSql.Show vbModeless
         Case "0202": FrmDescProj.Show vbModeless

      
'*****Case "03":SEGURANÇA
         Case "0300"
         Case "0301": 'Separador
         Case "0302"
            If UCase(ClsSeg.IDUSER) = "FIM" Then Exit Sub
            If SysMdi.Toolbar.AllowCustomize Then
              SysMdi.Toolbar.RestoreToolbar Sys.AppTitle, ANALISTA, SysMdi.Toolbar.Name
            End If
            With SysMdi.StatusBar
              .Panels(3).Text = IIf(Len(Trim(ANALISTA)) = 0, "DSVM", ANALISTA)
              .Panels(4).Text = Sys.NomeEmpresa
              .Panels(5).Text = "Rio de Janeiro"
            End With
            SysMdi.StatusBar.Refresh
            Call SetDefault(hWnd)

'******Case "04":JANELA
         Case "0400": Me.Arrange 0
      
'******Case "05":AJUDA
         Case "0500": 'Call AboutShow 'frmSobre.Show
         Case "0501": 'Call AboutShow 'frmSobre.Show
         Case "0502": 'Separator
         Case "0503": Call AboutShow(App) 'frmSobre.Show
     End Select
     
End Sub

Public Sub F_PROCURAR()
   If Not Sys.MDIFilho Is Nothing Then
      On Error Resume Next
      Sys.MDIFilho.F_PROCURAR
      Select Case Err
         Case 438:  On Error GoTo 0
      End Select
    End If
End Sub

Public Sub F_REFRESH()
   If Not Sys.MDIFilho Is Nothing Then
      On Error Resume Next
      Sys.MDIFilho.F_REFRESH
      Select Case Err
         Case 438:  On Error GoTo 0
      End Select
   End If
End Sub

Public Sub F_SALVAR()
   If Not Sys.MDIFilho Is Nothing Then
      On Error Resume Next
      Sys.MDIFilho.F_SALVAR
      Select Case Err
         Case 438:  On Error GoTo 0
      End Select
    End If
End Sub

Public Function F_VOLTAR()
   On Error Resume Next
   If Forms.Count = 1 Then
      Unload Me
   Else
      If Sys.MDIFilho.Name = "FrmCad" Then
         Sys.MDIFilho.Visible = False
      Else
         Unload Sys.MDIFilho
      End If
   End If
End Function
Private Sub MDIForm_Activate()
   Call SetHourglass(hWnd)
   Screen.MousePointer = vbHourglass
   If Flag_Inicio% Then
      
      With SysMdi
         .Caption = App.Title + " " + LoadMsg(49) + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor))
         .WindowState = vbMaximized
         .TmBD.Enabled = True
         .Visible = False
      End With
   '====voltar
      Call ToolbarPrincipal
      If Me.Toolbar.AllowCustomize Then
'         Me.Toolbar.RestoreToolbar Sys.AppName, ClsUser.ID, Me.Toolbar.Name
      End If
      With SysMdi.StatusBar
         .Panels(3).Text = "InterProject" 'IIf(Len(Trim(ClsUser.ID)) = 0, "DSVM", ClsUser.ID)
         .Panels(4).Text = "" 'SYS_Nome_Filial$
         .Panels(5).Text = "TecaPlus - Rio de Janeiro" 'MyCONFIG.CODLOCAL & " - " & MyCONFIG.DSCLOCAL
      End With
      Flag_Inicio% = False
      Me.Visible = True
   End If
   Me.Toolbar.Visible = True
'   SysMdi.Toolbar.Refresh
   Call MDIForm_Click
   Screen.MousePointer = vbDefault
   'Call SetDefault(hWnd)
End Sub

Private Sub MDIForm_Click()
   SysMdi.SetFocus
End Sub

Private Sub MDIForm_Load()
   Dim Titulo$

   Call ResizeForm(Me)
   Set SysMdi = MdiPrincipal
   
   '* Caption da aplicação
   Titulo$ = Titulo$ & IIf(Sys.MICRO = "DSVM", " [Base Case]", "")
   Dim ClsLoad As New Load
   Call SetHourglass(hWnd)
   With ClsLoad
      .AnoDsvm = "December - 1997"
      .Aplic = App
      
      .Show
      Sys.AppExe = .EXEName             '* RM
      Sys.AppName = .Name               '* RMStatus
      Sys.AppTitle = .Title             '* Requisition Material Status
      Sys.AppVer = .Versao              '* Versão 3.0
      Sys.AppDate = .GetAppDateStart     '* 30/09/99273 - 10:16:03.
      Sys.NomeEmpresa = .Empresa    '* Marítima Petróleo Engenharia LTDA.
      If .Ativa Then End             '* Testa se já existe uma cópia da aplicação rodando.
      If Not .SetFormat Then End     '* Define formato Data e número.
   End With
   'App.HelpFile = DIRETORIO_REDE$ + "materiais.hlp"
   
   '* Variáveis do Registro do Sistema
   Call GetConfig
   
   '* Cria Banco de Dados se ele não existir e se o banco em uso for ACCESS
'   While Not FileExists(dbDrive & dbName)
'      ClsUser.ID = ANALISTA
'      ClsUser.Name = NMANALISTA
'      ClsUser.Grp = eGrpUser.GRPANALISTA
'      FrmConfig.Show vbModal
      'If InArray(Val(dbVersion), Array(1, 8, 16, 32)) Then
      '   If Not FileExists(dbdrive & dbName) Then
      '      Call CriarBD(dbdrive & dbName)
      'End If
'   Wend

   '* Conectar Servidor
'   Call Db.SrvConecta(dbdrive, dbName, "", "", "", "")
'   ClsUser.dBase = Db
'   ClsSeg.dBase = Db
'   ClsSeg.CodSys = AppExe
'   If MICRO = "DSVM" Then
'      ClsUser.ID = ANALISTA$
'      ClsUser.Name = NMANALISTA$
'      ClsUser.Grp = GRPANALISTA
'      ClsUser.LOGON = ANALISTA
'      DataInicioSistema = Date
'   Else
'      ClsSeg.Show
'   End If
'   DoEvents
'   If ClsSeg.IDUSER = "FIM" Then
'      Call Db.DerrubarSistema
'      Screen.MousePointer = vbDefault
'      Set ClsSeg = Nothing
'      Set ClsUser = Nothing
'      End
'   End If
'   ClsUser.ID = IIf(ClsSeg.IDUSER = "", ClsUser.ID, ClsSeg.IDUSER)
   '* Recuperar Menu
'   Call ClsSeg.RecuperarMenu(SysMdi)
    Call Wait(3)
   

   'Salva mês, dia e hora de entrada no sistema
   Sys.AppTimeStamp = Sys.AppName + "-" + Sys.AppDate + " "
   
'   BANCO.dBase = Db

   Call BotaoIcon(Me)
   
   '* Variáveis de inicialização

   Flag_Inicio% = True
   MdiPrincipal.MnuMousePrincipal.Visible = False
   Call SetDefault(hWnd)
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim RESP%
   RESP% = ExibirPergunta(LoadMsg(30), Sys.AppName + " " + Mid(Sys.AppVer, InStr(Sys.AppVer, ".") - 1))
   Cancel = (RESP% = vbNo)
   Screen.MousePointer = vbDefault
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
   Dim Form As Variant
   Call SetHourglass(hWnd)
'   If Me.Toolbar.AllowCustomize And ClsUser.Grp <= StrZero(eGrpUser.GRPMASTER, 3) Then
'      Me.Toolbar.SaveToolbar Sys.AppName, ClsUser.ID, Me.Toolbar.Name
'   End If
   For Each Form In Forms
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
   Dim i%
   If Index > 6 Then
      For i = 1 To DB.dBases.Count
         Me.Mnu00(6 + i).Checked = False
      Next
      Me.Mnu00(Index).Checked = True
      DB.dBase = DB.dBases(Index - 6)
'      With Db ' Db.dBases(Index - 6)
'         Call Db.SrvConecta(.dbdrive, .dbName, .DSN, .UID, .PWD, .StrDATABASE)
'      End With
      BANCO.dBase = DB
      Call Me.F_REFRESH
   Else
      Call F_OPCAO("00" + StrZero(Index, 2))
   End If
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
Public Sub MnuMouse00_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 0
         FrmView.Tabela = FrmSeekCrg.LstTabela
         FrmView.Show vbModeless
      Case 1
         MnuMouse00(Index).Checked = Not MnuMouse00(Index).Checked
         Call FrmSeekCrg.MontarLstTabela
   End Select
   Screen.MousePointer = vbDefault
End Sub
Private Sub MnuMouse01_Click(Index As Integer)
   Dim Pos%, Pos2%, Sql$
   Dim StrWHERE$, Campo$
   
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 0
         Dim MyPrint As IMPRESSAO
         Set MyPrint = New IMPRESSAO
         With MyPrint
            .CryRpt = MdiPrincipal.RELAT
            .Rpt_Drive = App.Path & "\"
            .Titulo = FrmView.Tabela
            .Aplic = App
            .Idioma = Sys_Idioma
            Call .ImprimeGrid(FrmView.GrdView)
         End With
         Set MyPrint = Nothing
      Case 1
         Select Case Mid(UCase(MnuMouse01(Index).Caption), 1, 4)
            Case "&ORD" 'ENAÇÃO MÚLTIPLA"
               With FrmView
                  .OrderBy = SetOrderBy(FrmView.GrdView)
                  If .OrderBy <> "" Then
                     Pos = InStr(UCase(.rDataEVT.Sql), "ORDER BY")
                     Sql = .rDataEVT.Sql
                     If Pos > 0 Then Sql = Mid(.rDataEVT.Sql, 1, Pos - 1)
                     Sql = Sql + .OrderBy
                     Call .MontaGridLocal(Sql)
                  End If
               End With
            Case "&FIL" 'TRAR COLUNA"
               Pos = InStr(MnuMouse01(Index).Caption, "(")
               Pos2 = InStr(MnuMouse01(Index).Caption, ")")
               Campo = Mid(MnuMouse01(Index).Caption, Pos + 1, Pos2 - Pos - 1)
               StrWHERE = InputBox("Digite o Filtro para o Campo " & Campo & " : ", "&Filtrar Coluna", Campo)
               If StrWHERE <> "" And InStr(StrWHERE, "=") <> 0 Then
                  StrWHERE = StrReplace(StrWHERE, """", "'")
                  With FrmView
                     Pos = InStr(UCase(.rDataEVT.Sql), "WHERE")
                     Sql = .rDataEVT.Sql
                     If InStr(UCase(.rDataEVT.Sql), StrWHERE) = 0 Then
                        If Pos > 0 And InStr(UCase(.rDataEVT.Sql), StrWHERE) = 0 Then
                           Sql = Mid(.rDataEVT.Sql, 1, Pos + 5) & StrWHERE & " AND " & Mid(.rDataEVT.Sql, Pos + 5)
                        Else
                           Pos = InStr(UCase(.rDataEVT.Sql), " ORDER BY")
                           If Pos > 0 Then
                              Sql = Mid(Sql, 1, Pos - 1) & " WHERE " & StrWHERE & Mid(Sql, Pos)
                           Else
                              Sql = Sql & " WHERE " & StrWHERE
                           End If
                        End If
                     End If
                     Call .MontaGridLocal(Sql)
                  End With
               
               End If
         End Select
         
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub TmBD_Timer()
    DB.MinutosConexao = DB.MinutosConexao + 1
'    If Db.MinutosConexao = (TEMPO_MAXIMO_CONEXAO% + 1) Then
'       Call Db.DerrubarSistema
'       End
'    End If
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
   Dim Frm As Form, Ind%
   Ind% = Val(Mid(Button.Key, 2, 2))
   Select Case Ind%
      Case 1: Call F_OPCAO("000000")
      Case 2: Call F_OPCAO("0200")
      Case 3: Call F_OPCAO("0201")
      Case 4: Call F_OPCAO("0201")
      Case 5:
      Case 6:
      Case 7:
      '==============================
      '==============================
      Case 8: Call F_OPCAO("0005")
      Case 9:
      Case 10: Call F_OPCAO("0100")
      Case 11: Call F_OPCAO("0201")
      Case 12: Call F_OPCAO("0202")
      Case 13: Call F_OPCAO("0203")
      Case 14: Call F_VOLTAR
      Case 15: Unload Me
      Case 16:
      Case 17: Call F_OPCAO("0302")
      Case 18: 'Set Frm = FrmTools
      Case 19: 'FrmBackup.Show vbModal
      Case 20: Call F_OPCAO("0503")
      Case 21:
   End Select
End Sub

