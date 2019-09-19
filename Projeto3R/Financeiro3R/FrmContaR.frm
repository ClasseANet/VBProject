VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Begin VB.Form FrmContaR 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Fatura"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.FlatEdit TxtDTPREV 
      Height          =   315
      Left            =   6600
      TabIndex        =   16
      Top             =   1560
      Width           =   1275
      _Version        =   720898
      _ExtentX        =   2249
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   10
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtIDCONTA 
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   14737632
      BackColor       =   14737632
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.DateTimePicker CmbDTPREV 
      Height          =   345
      Left            =   6600
      TabIndex        =   17
      Top             =   1560
      Width           =   1575
      _Version        =   720898
      _ExtentX        =   2778
      _ExtentY        =   609
      _StockProps     =   68
      Format          =   1
      CurrentDate     =   40356.1843055556
   End
   Begin XtremeSuiteControls.PushButton CmdLov 
      Height          =   345
      Left            =   1200
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Pesquisar"
      Top             =   360
      Visible         =   0   'False
      Width           =   375
      _Version        =   720898
      _ExtentX        =   661
      _ExtentY        =   609
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmContaR.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit TxtHISTORICO 
      Height          =   315
      Left            =   1080
      TabIndex        =   19
      Top             =   2040
      Width           =   7065
      _Version        =   720898
      _ExtentX        =   12462
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   80
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5640
      TabIndex        =   21
      Top             =   2760
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Cancelar"
      ForeColor       =   192
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdOk 
      Default         =   -1  'True
      Height          =   375
      Left            =   7080
      TabIndex        =   20
      Top             =   2760
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Ok"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdExcluir 
      Height          =   375
      Left            =   120
      TabIndex        =   24
      ToolTipText     =   "Excluir"
      Top             =   2760
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Excluir"
      ForeColor       =   64
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmContaR.frx":0183
   End
   Begin XtremeSuiteControls.PushButton CmdSalvar 
      Height          =   375
      Left            =   1320
      TabIndex        =   23
      Top             =   2760
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Salvar"
      ForeColor       =   32768
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmContaR.frx":0C4D
   End
   Begin XtremeSuiteControls.FlatEdit TxtAtend 
      Height          =   315
      Left            =   6600
      TabIndex        =   11
      Top             =   960
      Width           =   1545
      _Version        =   720898
      _ExtentX        =   2725
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Enabled         =   0   'False
      Text            =   "88/88/8888 88:88"
      MaxLength       =   80
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtIDATEND 
      Height          =   315
      Left            =   5880
      TabIndex        =   10
      Top             =   960
      Width           =   735
      _Version        =   720898
      _ExtentX        =   1296
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   14737632
      Text            =   "888888"
      BackColor       =   14737632
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtValor 
      Height          =   345
      Left            =   6600
      TabIndex        =   4
      Top             =   360
      Width           =   1545
      _Version        =   720898
      _ExtentX        =   2725
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0,00"
      Alignment       =   1
      MaxLength       =   12
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GrpSessao 
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   5775
      _Version        =   720898
      _ExtentX        =   10186
      _ExtentY        =   1085
      _StockProps     =   79
      Appearance      =   4
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton CmdLovCli 
         Height          =   345
         Left            =   5400
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
         _Version        =   720898
         _ExtentX        =   661
         _ExtentY        =   609
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmContaR.frx":2517
      End
      Begin XtremeSuiteControls.FlatEdit TxtNOME 
         Height          =   345
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   4320
         _Version        =   720898
         _ExtentX        =   7620
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   16777215
         Text            =   "Patricia Moreira"
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdIDCLIENTE 
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   945
         _Version        =   720898
         _ExtentX        =   1667
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cliente"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   0
         Appearance      =   2
         ImageAlignment  =   6
         TextImageRelation=   0
      End
   End
   Begin XtremeSuiteControls.ComboBox CmbCategoria 
      Height          =   315
      Left            =   1080
      TabIndex        =   13
      Top             =   1560
      Width           =   2115
      _Version        =   720898
      _ExtentX        =   3731
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox CmbSubCategoria 
      Height          =   315
      Left            =   3360
      TabIndex        =   14
      Top             =   1560
      Width           =   2115
      _Version        =   720898
      _ExtentX        =   3731
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdVenda 
      Height          =   375
      Left            =   3120
      TabIndex        =   22
      Top             =   2760
      Width           =   1575
      _Version        =   720898
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Pagar"
      ForeColor       =   32768
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmContaR.frx":269A
   End
   Begin XtremeSuiteControls.Label LblVenc 
      Height          =   285
      Left            =   5640
      TabIndex        =   15
      Top             =   1560
      Width           =   960
      _Version        =   720898
      _ExtentX        =   1693
      _ExtentY        =   494
      _StockProps     =   79
      Caption         =   "Vencimento:"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label LblCategoria 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   960
      _Version        =   720898
      _ExtentX        =   1693
      _ExtentY        =   494
      _StockProps     =   79
      Caption         =   "Categoria:"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label LblAtend 
      Height          =   285
      Left            =   5880
      TabIndex        =   9
      Top             =   720
      Width           =   960
      _Version        =   720898
      _ExtentX        =   1693
      _ExtentY        =   494
      _StockProps     =   79
      Caption         =   "Atendimento:"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label LblObs 
      Height          =   285
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   960
      _Version        =   720898
      _ExtentX        =   1693
      _ExtentY        =   494
      _StockProps     =   79
      Caption         =   "Observação:"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label LblVALOR 
      Height          =   285
      Left            =   5880
      TabIndex        =   3
      Top             =   360
      Width           =   600
      _Version        =   720898
      _ExtentX        =   1058
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Valor:"
      Transparent     =   -1  'True
   End
   Begin VB.Label LblVenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº &Fatura"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "FrmContaR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://www.guaru.com.br/sistemas/document/pdvtef_06.asp
Option Explicit
Event Activate()
Event Load()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Unload(Cancel As Integer)
Event Resize()
Event CmdOkClick()
Event CmdCancelClick()
Event CmdExcluirClick()
Event CmdLovClick()
Event CmdLovCliClick()
Event CmdIDCLIENTEClick()

Event TxtIDCONTAGotFocus()
Event TxtIDCONTALostFocus()
Event TxtNOMEChange()
Event TxtNOMEKeyPress(KeyAscii As Integer)
Event TxtDTPREVLostFocus()

Private Sub LblAtend_Click()

End Sub

Private Sub TxtDTPREV_GotFocus()
   If Me.TxtDTPREV.Enabled Then
      Me.TxtDTPREV.SelStart = 0
      Me.TxtDTPREV.SelLength = Len(Me.TxtDTPREV.Text)
      Call SelecionarTexto(Me.TxtDTPREV)
   End If
End Sub
Private Sub TxtDTPREV_LostFocus()
   RaiseEvent TxtDTPREVLostFocus
End Sub
Private Sub CmdCancel_Click()
   RaiseEvent CmdCancelClick
End Sub
Private Sub CmdExcluir_Click()
   RaiseEvent CmdExcluirClick
End Sub
Private Sub CmdIDCLIENTE_Click()
   RaiseEvent CmdIDCLIENTEClick
End Sub
Private Sub CmdLov_Click()
   RaiseEvent CmdLovClick
End Sub
Private Sub CmdLovCli_Click()
   RaiseEvent CmdLovCliClick
End Sub
Private Sub CmdOk_Click()
   RaiseEvent CmdOkClick
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
Private Sub TxtIDCONTA_GotFocus()
   RaiseEvent TxtIDCONTAGotFocus
End Sub
Private Sub TxtIDCONTA_LostFocus()
   RaiseEvent TxtIDCONTALostFocus
End Sub
Private Sub TxtNOME_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtValor_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtValor_LostFocus()
   Me.TxtValor.Text = ValBr(Me.TxtValor.Text)
End Sub

