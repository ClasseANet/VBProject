VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmSenhaNova 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   2235
   ClientLeft      =   1155
   ClientTop       =   1230
   ClientWidth     =   3585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2235
   ScaleWidth      =   3585
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel1 
      Height          =   2235
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   3585
      _Version        =   65536
      _ExtentX        =   6324
      _ExtentY        =   3942
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BorderWidth     =   2
      Alignment       =   0
      Begin VB.CommandButton CmdOper 
         Caption         =   "&OK"
         Height          =   345
         Index           =   0
         Left            =   1360
         TabIndex        =   6
         Top             =   1770
         Width           =   1035
      End
      Begin VB.CommandButton CmdOper 
         Caption         =   "&Go Back"
         Height          =   345
         Index           =   1
         Left            =   2400
         TabIndex        =   7
         Top             =   1770
         Width           =   1035
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1515
         Left            =   150
         TabIndex        =   12
         Top             =   120
         Width           =   3285
         _Version        =   65536
         _ExtentX        =   5794
         _ExtentY        =   2672
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         BorderWidth     =   2
         BevelOuter      =   1
         Alignment       =   0
         Begin Threed.SSPanel PnlOld 
            Height          =   315
            Left            =   150
            TabIndex        =   0
            Top             =   180
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Old"
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Alignment       =   1
         End
         Begin Threed.SSPanel PnlNew 
            Height          =   315
            Left            =   150
            TabIndex        =   2
            Top             =   600
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "New"
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Alignment       =   1
         End
         Begin Threed.SSPanel PnlVerify 
            Height          =   315
            Left            =   150
            TabIndex        =   4
            Top             =   1020
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Verify"
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Alignment       =   1
         End
         Begin Threed.SSPanel SSPanel6 
            Height          =   315
            Left            =   1470
            TabIndex        =   8
            Top             =   180
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   15
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Alignment       =   1
            Autosize        =   3
            Begin VB.TextBox TxtSenha 
               ForeColor       =   &H00008000&
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   15
               MaxLength       =   64
               PasswordChar    =   "*"
               TabIndex        =   1
               Top             =   15
               Width           =   1635
            End
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Left            =   1470
            TabIndex        =   9
            Top             =   600
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   15
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Alignment       =   1
            Autosize        =   3
            Begin VB.TextBox TxtNovaSenha 
               ForeColor       =   &H00FF0000&
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   15
               MaxLength       =   64
               PasswordChar    =   "*"
               TabIndex        =   3
               Top             =   15
               Width           =   1635
            End
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   315
            Left            =   1470
            TabIndex        =   10
            Top             =   1020
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   15
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Alignment       =   1
            Autosize        =   3
            Begin VB.TextBox TxtConfirmacao 
               ForeColor       =   &H00FF0000&
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   15
               MaxLength       =   64
               PasswordChar    =   "*"
               TabIndex        =   5
               Top             =   15
               Width           =   1635
            End
         End
      End
   End
End
Attribute VB_Name = "FrmSenhaNova"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event CmdOperNew(Index As Integer)
Event ActiveSenhaNova()
Event LoadSenhaNova()
Event TxtSenhaLostFocus()
Private Sub CmdRetorna_Click()
    UnLoad Me
End Sub
Private Sub CmdOper_Click(Index As Integer)
   RaiseEvent CmdOperNew(Index)
End Sub
Private Sub Form_Activate()
   RaiseEvent ActiveSenhaNova
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case vbKeyEscape: UnLoad Me
      Case Else: KeyAscii = ClsDsr.SendTab(Me, KeyAscii)
   End Select
   DoEvents
End Sub
Private Sub Form_Load()
   RaiseEvent LoadSenhaNova
End Sub

Private Sub txtconfirmacao_GotFocus()
    Call ClsDsr.SelecionarTexto(ActiveControl)
End Sub
Private Sub txtNovasenha_GotFocus()
    Call ClsDsr.SelecionarTexto(ActiveControl)
End Sub
Private Sub TxtSenha_GotFocus()
    Call ClsDsr.SelecionarTexto(ActiveControl)
End Sub

Private Sub TxtSenha_LostFocus()
   RaiseEvent TxtSenhaLostFocus
End Sub
