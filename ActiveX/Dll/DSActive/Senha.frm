VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmSenha 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   2055
   ClientLeft      =   945
   ClientTop       =   1770
   ClientWidth     =   4305
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
   LinkTopic       =   "Senha"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2055
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel Pnl 
      Height          =   2085
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4320
      _Version        =   65536
      _ExtentX        =   7620
      _ExtentY        =   3678
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
      Begin VB.Timer Timer 
         Interval        =   60000
         Left            =   0
         Top             =   1560
      End
      Begin VB.CommandButton CmdNovaSenha 
         Caption         =   "&Change..."
         Height          =   345
         Left            =   2760
         TabIndex        =   5
         Top             =   1200
         Width           =   1365
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "&OK"
         Height          =   345
         Left            =   2780
         TabIndex        =   3
         Top             =   120
         Width           =   1365
      End
      Begin VB.CommandButton CmdReturn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Go Back"
         Height          =   345
         Left            =   2780
         TabIndex        =   4
         Top             =   480
         Width           =   1365
      End
      Begin Threed.SSPanel Pnl 
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   7
         Top             =   1680
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   450
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
         Alignment       =   0
      End
      Begin Threed.SSPanel Pnl 
         Height          =   1425
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   4471
         _ExtentY        =   2514
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
         Begin Threed.SSPanel PnlDt 
            Height          =   315
            Left            =   150
            TabIndex        =   9
            Top             =   150
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "&Date"
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
            Font3D          =   1
            Alignment       =   1
         End
         Begin Threed.SSPanel PnlUser 
            Height          =   315
            Left            =   150
            TabIndex        =   10
            Top             =   555
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "&User"
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
            Font3D          =   1
            Alignment       =   1
         End
         Begin Threed.SSPanel PnlPwd 
            Height          =   315
            Left            =   150
            TabIndex        =   11
            Top             =   960
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "&Password"
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
            Font3D          =   1
            Alignment       =   1
         End
         Begin Threed.SSPanel Pnl 
            Height          =   315
            Index           =   3
            Left            =   1080
            TabIndex        =   12
            Top             =   150
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
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
            Begin MSMask.MaskEdBox MskData 
               Height          =   285
               Left            =   15
               TabIndex        =   0
               Top             =   15
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   503
               _Version        =   393216
               ForeColor       =   32768
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
         End
         Begin Threed.SSPanel Pnl 
            Height          =   315
            Index           =   4
            Left            =   1080
            TabIndex        =   13
            Top             =   555
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
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
            Begin VB.TextBox TxtUsuario 
               ForeColor       =   &H00008000&
               Height          =   285
               Left            =   15
               MaxLength       =   10
               TabIndex        =   1
               Top             =   15
               Width           =   1275
            End
         End
         Begin Threed.SSPanel Pnl 
            Height          =   315
            Index           =   5
            Left            =   1080
            TabIndex        =   14
            Top             =   960
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
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
               MaxLength       =   10
               PasswordChar    =   "*"
               TabIndex        =   2
               Top             =   15
               Width           =   1275
            End
         End
      End
   End
End
Attribute VB_Name = "FrmSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event Load()
Event CmdOk()
Event CmdReturn()
Event CmdNovaSenha()
Event UnLoad()
Event Timer()
Private Sub CmdNovasenha_Click()
   RaiseEvent CmdNovaSenha
End Sub
Private Sub CmdOK_Click()
   RaiseEvent CmdOk
End Sub
Private Sub CmdReturn_Click()
   RaiseEvent CmdReturn
End Sub

Private Sub Form_Activate()
'   Dim r As Long
'   r = SetTopMostWindow(FrmSenha.hWnd, True)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case Else: KeyAscii = ClsDsr.SendTab(Me, KeyAscii)
   End Select
   DoEvents
End Sub
Private Sub Form_Load()
   Dim r As Long
   r = ClsAPI.SetTopMostWindow(FrmSenha.hWnd, True)
   RaiseEvent Load
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Dim r As Long
   r = ClsAPI.SetTopMostWindow(FrmSenha.hWnd, False)
   Timer.Tag = "0"
   RaiseEvent UnLoad
End Sub
Private Sub Timer_Timer()
   RaiseEvent Timer
End Sub

Private Sub TxtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ClsCtrl.Set_Focus(TxtSenha)
    Else
        KeyAscii = Asc(UCase(Chr$(KeyAscii)))
    End If
End Sub
Private Sub TxtSenha_Click()
    TxtSenha.Text = ""
End Sub
Private Sub TxtUsuario_GotFocus()
    Call ClsDsr.SelecionarTexto(ActiveControl)
End Sub
Private Sub TxtSenha_GotFocus()
    Call ClsDsr.SelecionarTexto(ActiveControl)
End Sub
Private Sub TxtSenha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0
        Call CmdOK_Click
    Else
        KeyAscii = Asc(UCase(Chr$(KeyAscii)))
    End If
End Sub
