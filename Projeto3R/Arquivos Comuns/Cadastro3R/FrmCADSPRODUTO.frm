VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Begin VB.Form FrmCADSPRODUTO 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Cadastro de Produtos"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GrpTipo 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   2655
      _Version        =   720898
      _ExtentX        =   4683
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   " Tipo  "
      Begin XtremeSuiteControls.RadioButton OptESERVICO 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
         _Version        =   720898
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Material"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton OptESERVICO 
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   975
         _Version        =   720898
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Servico"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit TxtID 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   705
      _Version        =   720898
      _ExtentX        =   1244
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   20
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtNMPROD 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   960
      Width           =   2415
      _Version        =   720898
      _ExtentX        =   4260
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   20
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdSair 
      Height          =   375
      Left            =   4680
      TabIndex        =   31
      Top             =   4800
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Sai&r"
      ForeColor       =   0
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdExcluir 
      Height          =   375
      Left            =   3000
      TabIndex        =   30
      Top             =   4800
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Excluir"
      ForeColor       =   192
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdSalvar 
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   4800
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Salvar"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdNovo 
      Height          =   375
      Left            =   1560
      TabIndex        =   29
      Top             =   4800
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Novo"
      ForeColor       =   4210752
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtCODPROD 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   705
      _Version        =   720898
      _ExtentX        =   1244
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   20
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GrpObjetivo 
      Height          =   735
      Left            =   3120
      TabIndex        =   11
      Top             =   1440
      Width           =   2775
      _Version        =   720898
      _ExtentX        =   4895
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   " Objetivo "
      Begin XtremeSuiteControls.RadioButton OptEVENDA 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   975
         _Version        =   720898
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Consumo"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton OptEVENDA 
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Width           =   975
         _Version        =   720898
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Venda"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.ComboBox CmbUNIDCONTROLE 
      Height          =   315
      Left            =   5040
      TabIndex        =   7
      Top             =   960
      Width           =   855
      _Version        =   720898
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GrpValores 
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   5775
      _Version        =   720898
      _ExtentX        =   10186
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   " Valores ($)"
      Begin XtremeSuiteControls.FlatEdit TxtVLVENDA 
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Top             =   720
         Width           =   825
         _Version        =   720898
         _ExtentX        =   1455
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Text            =   "33,33"
         Alignment       =   1
         MaxLength       =   20
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtVLCOMPRA 
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Top             =   240
         Width           =   825
         _Version        =   720898
         _ExtentX        =   1455
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "33,33"
         Alignment       =   1
         MaxLength       =   20
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtVLMEDIO 
         Height          =   285
         Left            =   3960
         TabIndex        =   20
         Top             =   240
         Width           =   825
         _Version        =   720898
         _ExtentX        =   1455
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   14737632
         Enabled         =   0   'False
         Text            =   "33,33"
         BackColor       =   14737632
         Alignment       =   1
         MaxLength       =   20
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   285
         Left            =   3240
         TabIndex        =   19
         Top             =   240
         Width           =   855
         _Version        =   720898
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Vl.Médio :"
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   855
         _Version        =   720898
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Vl.Compra :"
      End
      Begin XtremeSuiteControls.Label LblVLVENDA 
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   855
         _Version        =   720898
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Vl. Venda :"
         Enabled         =   0   'False
      End
   End
   Begin XtremeSuiteControls.GroupBox GrpEstoque 
      Height          =   1095
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   5775
      _Version        =   720898
      _ExtentX        =   10186
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   " Estoque (Qtd.)"
      Begin XtremeSuiteControls.FlatEdit TxtQtdMin 
         Height          =   285
         Left            =   1200
         TabIndex        =   23
         Top             =   240
         Width           =   825
         _Version        =   720898
         _ExtentX        =   1455
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "33,33"
         Alignment       =   1
         MaxLength       =   20
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtSldDisponivel 
         Height          =   285
         Left            =   3960
         TabIndex        =   25
         Top             =   240
         Width           =   825
         _Version        =   720898
         _ExtentX        =   1455
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   14737632
         Enabled         =   0   'False
         Text            =   "33,33"
         BackColor       =   14737632
         Alignment       =   1
         MaxLength       =   20
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtQtdCompra 
         Height          =   285
         Left            =   1200
         TabIndex        =   27
         Top             =   600
         Width           =   825
         _Version        =   720898
         _ExtentX        =   1455
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "33,33"
         Alignment       =   1
         MaxLength       =   20
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label9 
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Qtd. Compra :"
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Qtd. Mínima :"
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   285
         Left            =   2640
         TabIndex        =   24
         Top             =   240
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Saldo Disponível :"
      End
   End
   Begin XtremeSuiteControls.CheckBox ChkATIVO 
      Height          =   255
      Left            =   5160
      TabIndex        =   32
      Top             =   120
      Width           =   735
      _Version        =   720898
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Ativo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
      Value           =   1
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   285
      Left            =   3960
      TabIndex        =   6
      Top             =   960
      Width           =   1215
      _Version        =   720898
      _ExtentX        =   2143
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Unid.Controle: "
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   615
      _Version        =   720898
      _ExtentX        =   1085
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Código :"
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
      _Version        =   720898
      _ExtentX        =   1508
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Id.:"
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   855
      _Version        =   720898
      _ExtentX        =   1508
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Nome :"
   End
End
Attribute VB_Name = "FrmCADSPRODUTO"
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
Event TxtIDLostFocus()
Event TxtCODPRODLostFocus()
Event OptEVENDAClick(Index As Integer)
Event TxtVLCOMPRAKeyDown(KeyCode As Integer, Shift As Integer)
Event TxtVLCOMPRALostFocus()
Event TxtVLCOMPRAKeyPress(KeyAscii As Integer)
Event TxtVLVENDAKeyDown(KeyCode As Integer, Shift As Integer)
Event TxtVLVENDALostFocus()
Event TxtVLVENDAKeyPress(KeyAscii As Integer)
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

Private Sub FlatEdit1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub OptEVENDA_Click(Index As Integer)
   RaiseEvent OptEVENDAClick(Index)
End Sub
Private Sub TxtCODPROD_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtCODPROD_LostFocus()
   RaiseEvent TxtCODPRODLostFocus
End Sub
Private Sub TxtID_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtID_LostFocus()
   RaiseEvent TxtIDLostFocus
End Sub
Private Sub TxtNMPROD_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtQtdCompra_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtQtdMin_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtVLCOMPRA_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtVLCOMPRA_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent TxtVLCOMPRAKeyDown(KeyCode, Shift)
End Sub
Private Sub TxtVLCOMPRA_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtVLCOMPRAKeyPress(KeyAscii)
End Sub
Private Sub TxtVLCOMPRA_LostFocus()
   RaiseEvent TxtVLVENDALostFocus
End Sub
Private Sub TxtVLVENDA_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtVLVENDA_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent TxtVLVENDAKeyDown(KeyCode, Shift)
End Sub
Private Sub TxtVLVENDA_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtVLVENDAKeyPress(KeyAscii)
End Sub
Private Sub TxtVLVENDA_LostFocus()
   RaiseEvent TxtVLVENDALostFocus
End Sub
