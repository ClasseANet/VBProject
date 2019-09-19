VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Begin VB.Form FrmCADOMAQ 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Máquina IPL"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.FlatEdit TxtID 
      Height          =   330
      Left            =   660
      TabIndex        =   1
      Top             =   120
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   582
      _StockProps     =   77
      BackColor       =   -2147483643
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdLov 
      Height          =   375
      Left            =   1890
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   375
      _Version        =   720898
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCADOMAQ.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit TxtCODMAQUINA 
      Height          =   330
      Left            =   4080
      TabIndex        =   6
      Top             =   600
      Width           =   1680
      _Version        =   720898
      _ExtentX        =   2963
      _ExtentY        =   582
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   50
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton OptATIVO 
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   17
      Top             =   120
      Width           =   975
      _Version        =   720898
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Inativo"
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton OptATIVO 
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   16
      Top             =   120
      Width           =   975
      _Version        =   720898
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Ativo"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtDTOPERACAO 
      Height          =   330
      Left            =   1320
      TabIndex        =   8
      Top             =   1080
      Width           =   1665
      _Version        =   720898
      _ExtentX        =   2937
      _ExtentY        =   582
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "88/88/8888"
      MaxLength       =   14
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox CmbIDTPMAQ 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   1665
      _Version        =   720898
      _ExtentX        =   2937
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox CmbIDSALA 
      Height          =   315
      Left            =   4080
      TabIndex        =   10
      Top             =   1080
      Width           =   1665
      _Version        =   720898
      _ExtentX        =   2937
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   630
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   5775
      _Version        =   720898
      _ExtentX        =   10186
      _ExtentY        =   1111
      _StockProps     =   79
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   1
      Begin XtremeSuiteControls.PushButton CmdExcluir 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   180
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Excluir"
         ForeColor       =   64
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmCADOMAQ.frx":0183
      End
      Begin XtremeSuiteControls.PushButton CmdSair 
         Height          =   375
         Left            =   4440
         TabIndex        =   15
         Top             =   180
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Sair"
         ForeColor       =   0
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdNovo 
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Novo"
         ForeColor       =   32768
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmCADOMAQ.frx":0C4D
      End
      Begin XtremeSuiteControls.PushButton CmdSalvar 
         Height          =   375
         Left            =   3000
         TabIndex        =   14
         Top             =   180
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Salvar"
         ForeColor       =   32768
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmCADOMAQ.frx":0DA7
      End
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   330
      Left            =   3240
      TabIndex        =   9
      Top             =   1080
      Width           =   735
      _Version        =   720898
      _ExtentX        =   1296
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Sala"
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   330
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Data de Início"
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   735
      _Version        =   720898
      _ExtentX        =   1296
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Tipo"
   End
   Begin XtremeSuiteControls.Label Label 
      Height          =   330
      Left            =   3240
      TabIndex        =   5
      Top             =   600
      Width           =   735
      _Version        =   720898
      _ExtentX        =   1296
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Código"
   End
   Begin XtremeSuiteControls.Label LblId 
      Height          =   330
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   510
      _Version        =   720898
      _ExtentX        =   900
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Id.:"
   End
End
Attribute VB_Name = "FrmCADOMAQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event CmdSalvarClick()
Event CmdSairClick()
Event CmdNovoClick()
Event CmdExcluirClick()
Event CmbIDTPMAQClick()
Private Sub CmbIDTPMAQ_Click()
   RaiseEvent CmbIDTPMAQClick
End Sub
Private Sub CmdExcluir_Click()
   RaiseEvent CmdExcluirClick
End Sub
Private Sub CmdNovo_Click()
   RaiseEvent CmdNovoClick
End Sub
Private Sub CmdSair_Click()
   RaiseEvent CmdSairClick
End Sub
Private Sub CmdSalvar_Click()
   RaiseEvent CmdSalvarClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub

Private Sub TxtDTOPERACAO_LostFocus()
   TxtDTOPERACAO.Text = FormatarData(TxtDTOPERACAO.Text)
End Sub
Private Sub TxtCODMAQ_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub

