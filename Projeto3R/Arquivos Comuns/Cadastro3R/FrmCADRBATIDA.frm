VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Begin VB.Form FrmCADRBATIDA 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Registro de Horario"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6810
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   4800
   End
   Begin XtremeSuiteControls.PushButton CmdSair 
      Height          =   495
      Left            =   4680
      TabIndex        =   13
      Top             =   4680
      Width           =   1575
      _Version        =   720898
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Sair"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtCHAPA 
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
      _Version        =   720898
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "F00010"
      Alignment       =   2
      MaxLength       =   6
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtID 
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   2520
      Width           =   735
      _Version        =   720898
      _ExtentX        =   1296
      _ExtentY        =   873
      _StockProps     =   77
      ForeColor       =   4210752
      BackColor       =   -2147483643
      Enabled         =   0   'False
      Text            =   "000"
      Alignment       =   2
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtNOME 
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   2520
      Width           =   4095
      _Version        =   720898
      _ExtentX        =   7223
      _ExtentY        =   873
      _StockProps     =   77
      ForeColor       =   4210752
      BackColor       =   -2147483643
      Enabled         =   0   'False
      Text            =   "Nome do Calaborador"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtSENHA 
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      Top             =   4080
      Width           =   1335
      _Version        =   720898
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "F00010"
      PasswordChar    =   "*"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdSalvar 
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   4080
      Width           =   735
      _Version        =   720898
      _ExtentX        =   1296
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "&Ok"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox CmbSENTIDO 
      Height          =   480
      Left            =   1680
      TabIndex        =   8
      Top             =   3240
      Width           =   4935
      _Version        =   720898
      _ExtentX        =   8705
      _ExtentY        =   847
      _StockProps     =   77
      ForeColor       =   4210752
      BackColor       =   -2147483643
      Style           =   2
      UseVisualStyle  =   -1  'True
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox CmbUnidade 
      Height          =   480
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   4935
      _Version        =   720898
      _ExtentX        =   8705
      _ExtentY        =   847
      _StockProps     =   77
      ForeColor       =   4210752
      BackColor       =   -2147483643
      Enabled         =   0   'False
      Style           =   2
      UseVisualStyle  =   -1  'True
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   6495
      _Version        =   720898
      _ExtentX        =   11456
      _ExtentY        =   1720
      _StockProps     =   79
      Appearance      =   6
      Begin VB.PictureBox PctReg 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6000
         Picture         =   "FrmCADRBATIDA.frx":0000
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin XtremeSuiteControls.FlatEdit TxtHORA 
         Height          =   495
         Left            =   4560
         TabIndex        =   16
         Top             =   360
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   77
         ForeColor       =   16777215
         BackColor       =   49152
         Text            =   "88:88:88"
         BackColor       =   49152
         Alignment       =   2
         Locked          =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtData 
         Height          =   495
         Left            =   960
         TabIndex        =   17
         Top             =   360
         Width           =   3495
         _Version        =   720898
         _ExtentX        =   6165
         _ExtentY        =   873
         _StockProps     =   77
         Text            =   "88/88/88 - Quinta-Feria"
         BackColor       =   -2147483633
         Alignment       =   2
         Appearance      =   4
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   615
         _Version        =   720898
         _ExtentX        =   1085
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Data:"
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
      _Version        =   720898
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Unidade:"
      ForeColor       =   4210752
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   1335
      _Version        =   720898
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Sentido:"
      ForeColor       =   4210752
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   1335
      _Version        =   720898
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Senha:"
      ForeColor       =   4210752
   End
   Begin XtremeSuiteControls.Label LblNome 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
      _Version        =   720898
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Nome:"
      ForeColor       =   4210752
   End
   Begin XtremeSuiteControls.Label LblMat 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
      _Version        =   720898
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Matrícula:"
      ForeColor       =   4210752
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
      _Version        =   720898
      _ExtentX        =   1720
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Hora:"
      Alignment       =   1
   End
End
Attribute VB_Name = "FrmCADRBATIDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event CmdSairClick()
Event CmdSalvarClick()
Event CmbSENTIDOKeyPress(KeyAscii As Integer)
Event CmbUNIDADEKeyPress(KeyAscii As Integer)
Event TxtCHAPALostFocus()
Event TxtCHAPAKeyPress(KeyAscii As Integer)
Event TxtCHAPAGotFocus()
Event TxtSENHALostFocus()
Event TxtSENHAKeyPress(KeyAscii As Integer)
Event Timer()
Private Sub CmbSENTIDO_KeyPress(KeyAscii As Integer)
   RaiseEvent CmbSENTIDOKeyPress(KeyAscii)
End Sub
Private Sub CmbUnidade_KeyPress(KeyAscii As Integer)
   RaiseEvent CmbSENTIDOKeyPress(KeyAscii)
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
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Timer1_Timer()
   RaiseEvent Timer
End Sub
Private Sub TxtCHAPA_GotFocus()
   RaiseEvent TxtCHAPAGotFocus
End Sub
Private Sub TxtCHAPA_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtCHAPAKeyPress(KeyAscii)
End Sub
Private Sub TxtCHAPA_LostFocus()
   RaiseEvent TxtCHAPALostFocus
End Sub
Private Sub TxtSENHA_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtSENHA_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtSENHAKeyPress(KeyAscii)
End Sub
Private Sub TxtSENHA_LostFocus()
   RaiseEvent TxtSENHALostFocus
End Sub
