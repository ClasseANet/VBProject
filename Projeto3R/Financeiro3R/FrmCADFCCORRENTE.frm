VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.ShortcutBar.v11.2.2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Begin VB.Form FrmCADFCCORRENTE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   Caption         =   "Detalhes da Conta"
   ClientHeight    =   4950
   ClientLeft      =   6090
   ClientTop       =   1965
   ClientWidth     =   4515
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   WindowState     =   1  'Minimized
   Begin XtremeSuiteControls.CheckBox ChkEVENDA 
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   3960
      Width           =   2055
      _Version        =   720898
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Conta Recebe Venda?"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtIDCONTA 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   14737632
      BackColor       =   14737632
      Locked          =   -1  'True
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtDSCCONTA 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   3015
      _Version        =   720898
      _ExtentX        =   5318
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   16777215
      Text            =   "Patricia Moreira"
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtNUMBANCO 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   1680
      Width           =   3015
      _Version        =   720898
      _ExtentX        =   5318
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   16777215
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtNUMCONTA 
      Height          =   315
      Left            =   1200
      TabIndex        =   9
      Top             =   2640
      Width           =   2295
      _Version        =   720898
      _ExtentX        =   4048
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   16777215
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtDVCONTA 
      Height          =   315
      Left            =   3600
      TabIndex        =   10
      Top             =   2640
      Width           =   615
      _Version        =   720898
      _ExtentX        =   1085
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   16777215
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtNUMAGENCIA 
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   2160
      Width           =   3015
      _Version        =   720898
      _ExtentX        =   5318
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   16777215
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GrpTipo 
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   4095
      _Version        =   720898
      _ExtentX        =   7223
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   " Tipo de conta"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.RadioButton OptTPCONTA 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1215
         _Version        =   720898
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Banco"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton OptTPCONTA 
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   1575
         _Version        =   720898
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Dinheiro / Caixa"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton OptTPCONTA 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   1695
         _Version        =   720898
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cartão de Crédito"
         BackColor       =   -2147483633
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton OptTPCONTA 
         Height          =   375
         Index           =   3
         Left            =   1920
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
         _Version        =   720898
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Investimento"
         BackColor       =   -2147483633
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton CmdSalvar 
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   4440
      Width           =   1695
      _Version        =   720898
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Salvar"
      ForeColor       =   32768
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCADFCCORRENTE.frx":0000
   End
   Begin XtremeSuiteControls.CheckBox ChkATIVO 
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   3960
      Width           =   615
      _Version        =   720898
      _ExtentX        =   1085
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Ativo"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption SccConta 
      Height          =   420
      Left            =   0
      TabIndex        =   18
      Top             =   60
      Width           =   4800
      _Version        =   720898
      _ExtentX        =   8467
      _ExtentY        =   741
      _StockProps     =   14
      Caption         =   "Conta: Caixa"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num. Conta"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num. Agência"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num. Banco"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desc. Conta"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Id."
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   180
   End
End
Attribute VB_Name = "FrmCADFCCORRENTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Load()
Event Rezise()
Event Unload(Cancel As Integer)
Event CmdSalvarClick()
Private Sub CmdSalvar_Click()
   RaiseEvent CmdSalvarClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   RaiseEvent Load
End Sub
Private Sub Form_Resize()
   RaiseEvent Rezise
End Sub
Private Sub Form_Unload(Cancel As Integer)
   RaiseEvent Unload(Cancel)
End Sub

