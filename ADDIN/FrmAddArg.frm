VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmAddArg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adicionar Argumento"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   2100
   ClientWidth     =   3405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4485
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox TxtDefaultValue 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox ChkOptional 
      BackColor       =   &H00C0FFFF&
      Caption         =   "By Val"
      Height          =   200
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   200
   End
   Begin VB.ListBox LstTipo 
      Height          =   2205
      ItemData        =   "FrmAddArg.frx":0000
      Left            =   120
      List            =   "FrmAddArg.frx":0025
      TabIndex        =   7
      Top             =   960
      Width           =   3135
   End
   Begin VB.CheckBox ChkByVal 
      BackColor       =   &H00C0FFFF&
      Caption         =   "By Val"
      Height          =   200
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   200
   End
   Begin VB.TextBox TxtNome 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      MaxLength       =   20
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   3960
      WhatsThisHelpID =   10287
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Top             =   3960
      WhatsThisHelpID =   10289
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "&Valor Padrão"
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
      Index           =   4
      Left            =   1320
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Tipo de Dado"
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
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nome : "
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "&Optional"
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
      Index           =   3
      Left            =   480
      TabIndex        =   9
      Top             =   3600
      Width           =   2955
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "&By Val"
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
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   3360
      Width           =   2955
   End
End
Attribute VB_Name = "FrmAddArg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Load()
Event CmdOperClick(Index As Integer)
Event ChkOptionalClick()
Event TxtNomeKeyPress(KeyAscii As Integer)
Private Sub ChkOptional_Click()
   RaiseEvent ChkOptionalClick
End Sub

Private Sub CmdOper_Click(Index As Integer)
   RaiseEvent CmdOperClick(Index)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Form_Resize()
   Call PintarFundo(Me, Sys.Proj.FundoTela)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Call SetDefault(hwnd)
End Sub
Private Sub Lbl_Click(Index As Integer)
   Select Case Index
      Case 2: Me.ChkByVal.Value = IIf(Me.ChkByVal, vbUnchecked, vbChecked)
      Case 3: Me.ChkOptional.Value = IIf(Me.ChkOptional, vbUnchecked, vbChecked)
   End Select
End Sub
Private Sub LstTipo_DblClick()
   RaiseEvent CmdOperClick(0)
End Sub

Private Sub TxtNome_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub

Private Sub TxtNome_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtNomeKeyPress(KeyAscii)
End Sub
