VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmLoad 
   BorderStyle     =   0  'None
   ClientHeight    =   4200
   ClientLeft      =   1575
   ClientTop       =   1485
   ClientWidth     =   7410
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FontTransparent =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4200
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin Threed.SSPanel Pnl 
      Height          =   1110
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   1958
      _StockProps     =   15
      Caption         =   " Sistema"
      ForeColor       =   255
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
      BorderWidth     =   4
      Autosize        =   1
      MousePointer    =   11
   End
   Begin Threed.SSPanel Pnl 
      Height          =   4170
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7395
      _Version        =   65536
      _ExtentX        =   13044
      _ExtentY        =   7355
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
      BorderWidth     =   4
      MousePointer    =   11
      Begin VB.Frame Fre 
         Height          =   3300
         Left            =   360
         MousePointer    =   11  'Hourglass
         TabIndex        =   2
         Top             =   240
         Width           =   6705
         Begin VB.Label LblSistName 
            Alignment       =   2  'Center
            Caption         =   "Descrição do Sistema"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   120
            MousePointer    =   11  'Hourglass
            TabIndex        =   6
            Top             =   1380
            Width           =   6165
         End
         Begin VB.Label lblVersao 
            Alignment       =   2  'Center
            Caption         =   "Versão"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   120
            MousePointer    =   11  'Hourglass
            TabIndex        =   5
            Top             =   1740
            Width           =   6135
         End
         Begin VB.Label LblAno 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Mês - Ano"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5400
            MousePointer    =   11  'Hourglass
            TabIndex        =   4
            Top             =   3000
            Width           =   840
         End
         Begin VB.Label LblEmpresa 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Empresa Autorizada"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   105
            MousePointer    =   11  'Hourglass
            TabIndex        =   3
            Top             =   2970
            Width           =   1710
         End
      End
   End
End
Attribute VB_Name = "FrmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event Load()
Event Activate()
Private Sub Form_Activate()
   RaiseEvent Activate
'   Dim r As Long
'   r = SetTopMostWindow(FrmSenha.hWnd, True)
End Sub
Private Sub Form_Load()
   Dim r As Long
   r = ClsAPI.SetTopMostWindow(FrmLoad.hWnd, True)
   RaiseEvent Load
End Sub
