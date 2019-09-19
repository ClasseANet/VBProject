VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmOpBanco 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Banco de Dados"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   2100
   ClientWidth     =   5985
   ForeColor       =   &H00008000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3540
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin Threed.SSCommand CmdOper 
      Height          =   495
      Index           =   0
      Left            =   4560
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&OK"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin TabDlg.SSTab TabOpBanco 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Existente"
      TabPicture(0)   =   "FrmOpBanco.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "OptBanco(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "OptBanco(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Recente"
      TabPicture(1)   =   "FrmOpBanco.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "LstBdRecent"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.ListBox LstBdRecent 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   5295
      End
      Begin VB.OptionButton OptBanco 
         Caption         =   "MS-ACCESS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   300
         Index           =   0
         Left            =   -73800
         TabIndex        =   2
         Top             =   1080
         Value           =   -1  'True
         Width           =   2940
      End
      Begin VB.OptionButton OptBanco 
         Caption         =   "ODBC"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   300
         Index           =   1
         Left            =   -73800
         TabIndex        =   1
         Top             =   1800
         Width           =   2940
      End
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   495
      Index           =   1
      Left            =   3240
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin VB.Image ImgFundo 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Left            =   0
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "FrmOpBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdOper_Click(Index As Integer)
   If Index = 1 Then
      DB.CodeSql = 1
   Else
      If Me.TabOpBanco.Tab = 0 Then
         DB.isODBC = Me.OptBanco(1).Value
      Else
         If Me.LstBdRecent.List(Me.LstBdRecent.ListIndex) <> "" Then
            DB.Alias = Me.LstBdRecent.List(Me.LstBdRecent.ListIndex)
            If Not DB.isODBC Then
               DB.dbDrive = GetNameFromPath(Me.LstBdRecent.List(Me.LstBdRecent.ListIndex), 1)
               DB.dbName = GetNameFromPath(Me.LstBdRecent.List(Me.LstBdRecent.ListIndex), 2)
            Else
               DB.StrConect = Me.LstBdRecent.List(Me.LstBdRecent.ListIndex)
               Call DB.GetODBCConnectParts(Me.LstBdRecent.List(Me.LstBdRecent.ListIndex))
            End If
         End If
      End If
      'Call DB.SrvConecta(dbDrive, dbName, DSN, UID, PWD, StrDATABASE, Alias)
   End If
   Unload Me
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim cAux As String
   Call SetHourglass(hwnd)
   Me.TabOpBanco.Tab = 0
   LstBdRecent.Clear
   For i = 1 To 5
      cAux = Trim(GetSetting(Sys.Constru.AppName, "Outros", "BDRecente" & CStr(i), ""))
      If cAux <> "" Then Me.LstBdRecent.AddItem cAux
   Next
   Call ConfigForm(Me, SysMdi.Icon, Sys.Proj.FundoTela)
   Call SetDefault(hwnd)
End Sub
Private Sub LstBdRecent_DblClick()
   Call CmdOper_Click(0)
End Sub
