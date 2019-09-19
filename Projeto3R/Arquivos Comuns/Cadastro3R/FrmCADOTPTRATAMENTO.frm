VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Begin VB.Form FrmCADOTPTRATAMENTO 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Cadastro de Tratamentos"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.CheckBox ChkProdVenda 
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   1320
      Width           =   1575
      _Version        =   720898
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Produto de Venda"
      ForeColor       =   16711680
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   5775
      _Version        =   720898
      _ExtentX        =   10186
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   " Controles "
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox ChkFLGDISPARO 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2175
         _Version        =   720898
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Equipamento"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox ChkFLGAREA 
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   240
         Width           =   855
         _Version        =   720898
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Área"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox ChkFLGAVALIACAO 
         Height          =   255
         Left            =   4560
         TabIndex        =   20
         Top             =   240
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Avaliação"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton CmdPadrao 
      Height          =   345
      Left            =   5600
      TabIndex        =   8
      ToolTipText     =   "Cor Original"
      Top             =   885
      Width           =   375
      _Version        =   720898
      _ExtentX        =   661
      _ExtentY        =   609
      _StockProps     =   79
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCADOTPTRATAMENTO.frx":0000
   End
   Begin Cadastro3R.ctrlThemeColor CtlColor 
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
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
   Begin XtremeSuiteControls.FlatEdit TxtCampo01 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   4905
      _Version        =   720898
      _ExtentX        =   8652
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   30
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdSair 
      Height          =   375
      Left            =   4920
      TabIndex        =   24
      Top             =   3240
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
      Left            =   3240
      TabIndex        =   23
      Top             =   3240
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
      Left            =   120
      TabIndex        =   21
      Top             =   3240
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
      TabIndex        =   22
      Top             =   3240
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Novo"
      ForeColor       =   4210752
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtCampo02 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   705
      _Version        =   720898
      _ExtentX        =   1244
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   20
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox ChkATIVO 
      Height          =   255
      Left            =   5160
      TabIndex        =   25
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
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   5775
      _Version        =   720898
      _ExtentX        =   10186
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   " Produto de Venda"
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit TxtNMPROD 
         Height          =   285
         Left            =   840
         TabIndex        =   13
         Top             =   480
         Width           =   3360
         _Version        =   720898
         _ExtentX        =   5927
         _ExtentY        =   503
         _StockProps     =   77
         ForeColor       =   -2147483631
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         MaxLength       =   30
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdLovProd 
         Height          =   315
         Left            =   4200
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   465
         Width           =   375
         _Version        =   720898
         _ExtentX        =   661
         _ExtentY        =   556
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmCADOTPTRATAMENTO.frx":039A
      End
      Begin XtremeSuiteControls.FlatEdit TxtVLVENDA 
         Height          =   285
         Left            =   4800
         TabIndex        =   16
         Top             =   480
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
      Begin XtremeSuiteControls.FlatEdit TxtCodProd 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   705
         _Version        =   720898
         _ExtentX        =   1244
         _ExtentY        =   503
         _StockProps     =   77
         ForeColor       =   -2147483631
         BackColor       =   -2147483643
         Enabled         =   0   'False
         MaxLength       =   20
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   855
         _Version        =   720898
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Produto"
         ForeColor       =   -2147483631
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
         _Version        =   720898
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Cod"
         ForeColor       =   -2147483631
      End
      Begin XtremeSuiteControls.Label LblVLVENDA 
         Height          =   285
         Left            =   4800
         TabIndex        =   15
         Top             =   240
         Width           =   855
         _Version        =   720898
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Valor ($)"
         ForeColor       =   -2147483631
      End
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   840
      Width           =   1575
      _Version        =   720898
      _ExtentX        =   2778
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Cor de Apresentação :"
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   855
      _Version        =   720898
      _ExtentX        =   1508
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Frequência :"
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
      TabIndex        =   2
      Top             =   480
      Width           =   855
      _Version        =   720898
      _ExtentX        =   1508
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Descrição :"
   End
End
Attribute VB_Name = "FrmCADOTPTRATAMENTO"
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
Event CmdPadraoClick()
Event CmdLovProdClick()
Event ChkProdVendaClick()
Event TxtCampo01Change()
Event TxtCODPRODLostFocus()
Event TxtIDLostFocus()
Private Sub ChkProdVenda_Click()
   RaiseEvent ChkProdVendaClick
End Sub
Private Sub CmdExcluir_Click()
   RaiseEvent CmdExcluirClick
End Sub

Private Sub CmdLovProd_Click()
   RaiseEvent CmdLovProdClick
End Sub

Private Sub CmdNovo_Click()
   RaiseEvent CmdNovoClick
End Sub

Private Sub CmdPadrao_Click()
   RaiseEvent CmdPadraoClick
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
Private Sub TxtCampo01_Change()
   RaiseEvent TxtCampo01Change
End Sub
Private Sub TxtCampo01_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
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
Private Sub TxtVLVENDA_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
