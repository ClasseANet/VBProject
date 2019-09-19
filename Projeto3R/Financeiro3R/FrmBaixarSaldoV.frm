VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Begin VB.Form FrmBaixarSaldoV 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Baixar Saldo de Venda"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GrpSessao 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3495
      _Version        =   720898
      _ExtentX        =   6165
      _ExtentY        =   1296
      _StockProps     =   79
      Appearance      =   4
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit TxtNOME 
         Height          =   345
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   2895
         _Version        =   720898
         _ExtentX        =   5106
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   16777215
         Enabled         =   0   'False
         Text            =   "Patricia Moreira"
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtIDCLIENTE 
         Height          =   345
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   495
         _Version        =   720898
         _ExtentX        =   873
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   16777215
         Enabled         =   0   'False
         Text            =   "8888"
         BackColor       =   16777215
         Alignment       =   2
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label LblTel2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cliente"
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin XtremeSuiteControls.FlatEdit TxtIDVENDA 
      Height          =   345
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   16777215
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtDTVENDA 
      Height          =   345
      Left            =   2520
      TabIndex        =   6
      Top             =   360
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   16777215
      Enabled         =   0   'False
      Text            =   "01/01/2000"
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdSair 
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Sai&r"
      ForeColor       =   0
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdOk 
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Ok"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtSaldoNovo 
      Height          =   345
      Left            =   2520
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      BackColor       =   16777215
      Alignment       =   1
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtSaldo 
      Height          =   345
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Text            =   "5"
      BackColor       =   16777215
      Alignment       =   1
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Novo Saldo "
      Height          =   195
      Left            =   2520
      TabIndex        =   13
      Top             =   1680
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Atual"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      Height          =   195
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   345
   End
   Begin VB.Label LblVenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Venda"
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "FrmBaixarSaldoV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event CmdOkClick()
Event CmdSairClick()
Event TxtIDVENDALostFocus()
Event TxtSaldoNovoLostFocus()
Private Sub CmdSair_Click()
   RaiseEvent CmdSairClick
End Sub
Private Sub CmdOk_Click()
   RaiseEvent CmdOkClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub TxtIDVENDA_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtIDVENDA_LostFocus()
   RaiseEvent TxtIDVENDALostFocus
End Sub
Private Sub TxtSaldoNovo_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub

Private Sub TxtSaldoNovo_LostFocus()
   RaiseEvent TxtSaldoNovoLostFocus
End Sub
