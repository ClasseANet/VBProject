VERSION 5.00
Object = "{2401AC19-8566-4347-9B14-31C80AA9AEF0}#8.0#0"; "MCIControl.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Begin VB.Form FrmRegistry 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3600
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraVerif 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Verificação"
      Height          =   3255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3375
      Begin VB.Image ImgOK 
         Height          =   240
         Left            =   360
         Picture         =   "FrmRegistry.frx":0000
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImgNOK 
         Height          =   240
         Left            =   360
         Picture         =   "FrmRegistry.frx":18BA
         Top             =   360
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label LblVerif 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Autenticação do Sistema"
         Height          =   195
         Index           =   7
         Left            =   720
         TabIndex        =   16
         Top             =   2880
         Width           =   1770
      End
      Begin VB.Label LblVerif 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acesso ao Servidor de Dados"
         Height          =   195
         Index           =   6
         Left            =   720
         TabIndex        =   15
         Top             =   2520
         Width           =   2115
      End
      Begin VB.Label LblVerif 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acesso ao Banco de Dados"
         Height          =   195
         Index           =   5
         Left            =   720
         TabIndex        =   14
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label LblVerif 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acesso ao Servidor de Banco"
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   13
         Top             =   1800
         Width           =   2115
      End
      Begin VB.Label LblVerif 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acesso ao Registro do Sistema"
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   12
         Top             =   1440
         Width           =   2205
      End
      Begin VB.Label LblVerif 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acesso a Pasta de Sistemas"
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   11
         Top             =   1080
         Width           =   2010
      End
      Begin VB.Label LblVerif 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acesso a Pasta de Arquivos"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   10
         Top             =   720
         Width           =   1995
      End
      Begin VB.Label LblVerif 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acesso ao FTP"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.ComboBox CmbDomainName 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox TxtPassword 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "#"
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   4080
      Width           =   2415
   End
   Begin VB.TextBox TxtUserName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Text            =   "dsr"
      Top             =   3720
      Width           =   2415
   End
   Begin MCIControls.MCIButton CmdRegistro 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   5040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Registro"
      ForeColor       =   8388608
      ForeHover       =   8388608
   End
   Begin MCIControls.MCIButton CmdSair 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   5040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Sai&r"
      ForeColor       =   -2147483642
      ForeHover       =   -2147483642
   End
   Begin XtremeSuiteControls.ProgressBar PrbFlood 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3380
      Visible         =   0   'False
      Width           =   3375
      _Version        =   720898
      _ExtentX        =   5953
      _ExtentY        =   450
      _StockProps     =   93
      Scrolling       =   2
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label LblDominio 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&DOMÍNIO : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   1005
   End
   Begin VB.Label Lblsenha 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&SENHA : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   825
   End
   Begin VB.Label LblUsuario 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&USUARIO : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   1035
   End
End
Attribute VB_Name = "FrmRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event Activate()
Event Load()
Event CmdSair()
Event CmdRegistroClick()
Private Sub CmdRegistro_Click()
   RaiseEvent CmdRegistroClick
End Sub
Private Sub CmdSair_Click()
   RaiseEvent CmdSair
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      SendKeys "{TAB}"
   End If
End Sub

Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub TxtPassword_GotFocus()
   Me.TxtPassword.SelStart = 0
   Me.TxtPassword.SelLength = Len(Me.TxtPassword.Text)
End Sub
Private Sub TxtUserName_GotFocus()
   Me.TxtUserName.SelStart = 0
   Me.TxtUserName.SelLength = Len(Me.TxtUserName.Text)
End Sub
