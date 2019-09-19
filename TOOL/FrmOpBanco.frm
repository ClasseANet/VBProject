VERSION 5.00
Begin VB.Form FrmOpBanco 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Banco de Dados"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   2100
   ClientWidth     =   7860
   ForeColor       =   &H00008000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6525
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdOper 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "CONECTAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   20
      Top             =   6000
      WhatsThisHelpID =   10527
      Width           =   1635
   End
   Begin VB.Frame Frme 
      Caption         =   "Database Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   3495
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.TextBox TxtAlias 
         Height          =   330
         Left            =   2040
         LinkTimeout     =   30
         TabIndex        =   22
         Top             =   600
         WhatsThisHelpID =   10536
         Width           =   1995
      End
      Begin VB.ComboBox CmbVersao 
         Height          =   315
         ItemData        =   "FrmOpBanco.frx":0000
         Left            =   2040
         List            =   "FrmOpBanco.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         WhatsThisHelpID =   10532
         Width           =   2025
      End
      Begin VB.Frame Frme 
         Caption         =   "ODBC"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Index           =   0
         Left            =   4440
         TabIndex        =   4
         Top             =   480
         WhatsThisHelpID =   10533
         Width           =   2055
         Begin VB.OptionButton OptODBC 
            Caption         =   "Sim"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   320
            WhatsThisHelpID =   10534
            Width           =   735
         End
         Begin VB.OptionButton OptODBC 
            Caption         =   "Não"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   6
            Top             =   320
            Value           =   -1  'True
            WhatsThisHelpID =   10535
            Width           =   735
         End
      End
      Begin VB.TextBox TxtDbName 
         Height          =   330
         Left            =   2040
         LinkTimeout     =   30
         TabIndex        =   10
         Top             =   1680
         WhatsThisHelpID =   10536
         Width           =   4875
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
         Index           =   0
         Left            =   6960
         TabIndex        =   21
         Top             =   1320
         WhatsThisHelpID =   10527
         Width           =   435
      End
      Begin VB.TextBox TxtDSN 
         Height          =   330
         Left            =   2040
         LinkTimeout     =   30
         TabIndex        =   12
         Top             =   2040
         WhatsThisHelpID =   10536
         Width           =   4875
      End
      Begin VB.TextBox TxtUID 
         Height          =   330
         Left            =   2040
         LinkTimeout     =   30
         TabIndex        =   14
         Top             =   2400
         WhatsThisHelpID =   10536
         Width           =   4875
      End
      Begin VB.TextBox TxtPWD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2040
         LinkTimeout     =   30
         PasswordChar    =   "#"
         TabIndex        =   16
         Top             =   2760
         WhatsThisHelpID =   10536
         Width           =   4875
      End
      Begin VB.CommandButton CmdDSN 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   " Registra DSN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   5520
         TabIndex        =   17
         Top             =   3120
         WhatsThisHelpID =   10527
         Width           =   1395
      End
      Begin VB.ComboBox CmbDbDrive 
         Height          =   315
         ItemData        =   "FrmOpBanco.frx":0069
         Left            =   2040
         List            =   "FrmOpBanco.frx":006B
         TabIndex        =   8
         Text            =   "TxtDbDrive"
         Top             =   1320
         WhatsThisHelpID =   10532
         Width           =   4875
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Versão do Banco"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   960
         WhatsThisHelpID =   10537
         Width           =   1665
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Servidor"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         WhatsThisHelpID =   10538
         Width           =   840
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   1
         Top             =   600
         WhatsThisHelpID =   10537
         Width           =   585
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Banco de Dados"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         WhatsThisHelpID =   10538
         Width           =   1590
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Data Source Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         WhatsThisHelpID =   10538
         Width           =   1815
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Usuário"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         WhatsThisHelpID =   10538
         Width           =   765
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Senha"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         WhatsThisHelpID =   10538
         Width           =   585
      End
   End
   Begin VB.CommandButton CmdDefault 
      Caption         =   "Padrão"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3600
      Width           =   840
   End
   Begin VB.ListBox LstBdRecent 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   120
      TabIndex        =   19
      Top             =   3960
      Width           =   7575
   End
End
Attribute VB_Name = "FrmOpBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmbConexao_Change()

End Sub

Private Sub CmbDbDrive_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub

Private Sub CmbDbDrive_LostFocus()
   If Trim(Me.CmbDbDrive.Text) <> "" Then
      If -1 = LocalizarCombo(Me.CmbDbDrive, Me.CmbDbDrive.Text, False) Then
         Me.CmbDbDrive.AddItem Me.CmbDbDrive.Text
      End If
   End If
End Sub

Private Sub CmdDefault_Click()
   Call LocalizarCombo(Me.CmbDbDrive, "TAMOIO")
   Me.TxtDbName.Text = "MAUA"
   Me.TxtUID.Text = "USU_VERIF"
   Me.TxtPWD.Text = "DIPLOMATA"
End Sub

Private Sub CmdOper_Click()
   Dim i As Integer
   If Trim(Me.TxtDbName.Text) = "" Then
      Call ExibirAviso("Campo Inválido", "")
      Me.TxtDbName.SetFocus
      Exit Sub
   End If

   XDb.isODBC = Me.OptODBC(0).Value
   
   '***************
   '* Conexão ADO *
   '***************
   With XDb
      .isODBC = False
      .isADO = True
      .Alias = Me.TxtAlias.Text
      .dbTipo = eDbTipo.SQL_SERVER
      .dbVersao = "7.0"
      Select Case .dbTipo
         Case eDbTipo.Access
'               .dbDrive = GetSetting(Sys.AppName, "Database Drive", "DBDRIVE", "C:\DSR\" + UCase(Sys.AppName) + "\")
'               .dbName = GetSetting(Sys.AppName, "Database Format", "DBNAME", UCase(Sys.AppExeName) & ".mdb")
         Case eDbTipo.SQL_SERVER
            .Server = Me.CmbDbDrive.Text
            .dbName = Me.TxtDbName.Text
         Case eDbTipo.ORACLE
'               .Server = GetSetting(Sys.AppName, "Database Format", "SERVER", "SERVIDOR_SQL")
'               .dbName = GetSetting(Sys.AppName, "Database Format", "DBNAME", UCase(Sys.AppName))
      End Select
      .DSN = ""
      .UID = "USU_VERIF"
      .PWD = "DIPLOMATA"
      If Trim(Me.TxtUID.Text) <> "" Then
         .UID = Me.TxtUID.Text
      End If
      If Trim(Me.TxtPWD.Text) <> "" Then
         .PWD = Me.TxtPWD.Text
      End If
      .Alias = IIf(.Alias = "", .dbName, .Alias)
   End With
      
   i = Mid(GetSecao(XDb.Alias), Len(SecaoBase) + 1)
   If i <= 0 Then
      Call RenumerarSetting
      Call ObjetoParaRegistro(XDb, 1)
   Else
      Call ObjetoParaRegistro(XDb, i)
   End If
   
      
   Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   KeyAscii = SendTab(Me, KeyAscii)
End Sub
Private Sub Form_Load()
   Dim i As Integer
   Dim cAux As String
   
   Call SetHourglass(hWnd)
   
   LstBdRecent.Clear
   CmbDbDrive.Clear
   
   Dim MyXdb As Object
   Dim sAlias  As String
   i = 1
   sAlias = Trim(ReadIniFile(mvarLocalReg, SecaoBase & i, "ALIAS", ""))
   While sAlias <> ""
      Set MyXdb = Nothing
      Set MyXdb = CreateObject("XBANCO01.DS_BANCO")
      Call RegistroParaObjeto(MyXdb, sAlias)
      Me.LstBdRecent.AddItem MyXdb.Alias
      CmbDbDrive.AddItem IIf(MyXdb.dbDrive = "", MyXdb.Server, MyXdb.dbDrive)
      i = i + 1
      sAlias = Trim(ReadIniFile(mvarLocalReg, SecaoBase & i, "Alias", ""))
   Wend
   
   Call ConfigForm(Me, Me.Icon)
   Call SetDefault(hWnd)
End Sub
Private Sub LstBdRecent_Click()
   Call RegistroParaObjeto(XDb, Me.LstBdRecent)
   
   Me.TxtAlias.Text = XDb.Alias
   Me.CmbDbDrive.Text = IIf(XDb.dbDrive = "", XDb.Server, XDb.dbDrive)
   If UCase(Mid(Me.CmbDbDrive.Text, 1, 7)) = "[LOCAL]" Then
      Me.CmbDbDrive.Text = "[Local]" & Mid(Me.CmbDbDrive.Text, 8)
   End If
   If UCase(Mid(Me.CmbDbDrive.Text, 1, Len(Environ("COMPUTERNAME")))) = UCase(Environ("COMPUTERNAME")) Then
      Me.CmbDbDrive.Text = "[Local]" & Mid(Me.CmbDbDrive.Text, Len(Environ("COMPUTERNAME")) + 1)
   End If
   If UCase(Mid(Me.CmbDbDrive.Text, 1, 8)) = "[REMOTE]" Then
      Me.CmbDbDrive.Text = "[Remote]" & Mid(Me.CmbDbDrive.Text, 9)
   End If

   Me.TxtDbName.Text = XDb.dbName
   Me.TxtDSN.Text = XDb.DSN
   Me.TxtPWD.Text = XDb.PWD
   Me.TxtUID.Text = XDb.UID
   Call LocalizarCombo(Me.CmbVersao, XDb.dbVersao)
End Sub

Private Sub LstBdRecent_DblClick()
   Call LstBdRecent_Click
   Call CmdOper_Click
End Sub

Private Sub TabOpBanco_DblClick()

End Sub
