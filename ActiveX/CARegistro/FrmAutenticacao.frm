VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Begin VB.Form FrmAutenticacao 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autenticação"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3285
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "FrmAutenticacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   3285
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton CmdAutenticacao 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Autenticação"
      ForeColor       =   16711680
      UseVisualStyle  =   -1  'True
      Appearance      =   6
   End
   Begin VB.Frame FraVerif 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Verificação"
      Height          =   3255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3015
      Begin VB.Image ImgVerif 
         Height          =   240
         Left            =   120
         Picture         =   "FrmAutenticacao.frx":0442
         Stretch         =   -1  'True
         Top             =   1080
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImgOK 
         Height          =   240
         Left            =   120
         Picture         =   "FrmAutenticacao.frx":0884
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImgNOK 
         Height          =   240
         Left            =   120
         Picture         =   "FrmAutenticacao.frx":213E
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
         Left            =   480
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
         Left            =   480
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
         Left            =   480
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
         Left            =   480
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
         Left            =   480
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
         Left            =   480
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
         Left            =   480
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
         Left            =   480
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.ComboBox CmbDomainName 
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Text            =   "Domain Name"
      Top             =   4440
      Width           =   2055
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
      Width           =   2055
   End
   Begin VB.TextBox TxtUserName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Text            =   "dsr"
      Top             =   3720
      Width           =   2055
   End
   Begin XtremeSuiteControls.ProgressBar PrbFlood 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3375
      Visible         =   0   'False
      Width           =   3015
      _Version        =   720898
      _ExtentX        =   5318
      _ExtentY        =   450
      _StockProps     =   93
      Scrolling       =   2
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdSair 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   4920
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Sai&r"
      ForeColor       =   192
      UseVisualStyle  =   -1  'True
      Appearance      =   6
   End
   Begin VB.Label LblDominio 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&DOMÍNIO : "
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   840
   End
   Begin VB.Label Lblsenha 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&SENHA : "
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   690
   End
   Begin VB.Label LblUsuario 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&USUARIO : "
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   870
   End
End
Attribute VB_Name = "FrmAutenticacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event Activate()
Event Load()
Event CmdSair()
Event CmdAutenticacaoClick()
Private Sub CmbDomainName_KeyPress(KeyAscii As Integer)
   Call KeyPress(KeyAscii)
End Sub
Private Sub CmdAutenticacao_Click()
   RaiseEvent CmdAutenticacaoClick
End Sub
Private Sub CmdAutenticacao0_Click()
   RaiseEvent CmdAutenticacaoClick
End Sub
Private Sub CmdSair_Click()
   RaiseEvent CmdSair
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub TxtPassword_GotFocus()
   Me.TxtPassword.SelStart = 0
   Me.TxtPassword.SelLength = Len(Me.TxtPassword.Text)
End Sub
Private Sub TxtPassword_KeyPress(KeyAscii As Integer)
   Call KeyPress(KeyAscii)
End Sub
Private Sub TxtUserName_GotFocus()
   Me.TxtUserName.SelStart = 0
   Me.TxtUserName.SelLength = Len(Me.TxtUserName.Text)
End Sub
Private Sub TxtUserName_KeyPress(KeyAscii As Integer)
   Call KeyPress(KeyAscii)
End Sub
Private Sub KeyPress(ByRef KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{TAB}"
      If Trim(Me.TxtUserName.Text) <> "" And Trim(Me.TxtPassword.Text) <> "" Then
         Call CmdAutenticacao_Click
      End If
   End If
End Sub
