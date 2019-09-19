VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Begin VB.Form FrmCADFDESPESA 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Cadastro de Despesas/Receitas"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.FlatEdit TxtID 
      Height          =   285
      Left            =   1320
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
      Left            =   1320
      TabIndex        =   3
      Top             =   480
      Width           =   5505
      _Version        =   720898
      _ExtentX        =   9710
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   20
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdSair 
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   1560
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
      Left            =   3600
      TabIndex        =   6
      Top             =   1560
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
      Left            =   840
      TabIndex        =   4
      Top             =   1560
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
      Left            =   2160
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Novo"
      ForeColor       =   4210752
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox CmbTPDESP 
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Top             =   960
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox CmbIDPAI 
      Height          =   315
      Left            =   4080
      TabIndex        =   11
      Top             =   960
      Width           =   2775
      _Version        =   720898
      _ExtentX        =   4895
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   285
      Left            =   2880
      TabIndex        =   9
      Top             =   960
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Nível Superior :"
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   975
      _Version        =   720898
      _ExtentX        =   1720
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Tipo :"
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
Attribute VB_Name = "FrmCADFDESPESA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event CmbTPDESPClick()
Event CmdSalvarClick()
Event CmdSairClick()
Event CmdNovoClick()
Event CmdExcluirClick()
Event TxtIDLostFocus()
Private Sub CmbTPDESP_Click()
   RaiseEvent CmbTPDESPClick
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
Private Sub TxtCampo01_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtCODPAI_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtID_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtID_LostFocus()
   RaiseEvent TxtIDLostFocus
End Sub
