VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Begin VB.Form FrmCADFCCORRENTE 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Cadastro de Contas"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Width           =   4545
      _Version        =   720898
      _ExtentX        =   8017
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   20
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdSair 
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   1080
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
      Left            =   2880
      TabIndex        =   6
      Top             =   1080
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
      TabIndex        =   4
      Top             =   1080
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
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Novo"
      ForeColor       =   4210752
      UseVisualStyle  =   -1  'True
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
      Caption         =   "Descri��o :"
   End
End
Attribute VB_Name = "FrmCADFCCORRENTE"
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
Private Sub TxtID_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub

Private Sub TxtID_LostFocus()
   RaiseEvent TxtIDLostFocus
End Sub
