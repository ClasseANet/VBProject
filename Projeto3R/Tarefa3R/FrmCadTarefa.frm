VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Begin VB.Form FrmCadTarefa 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tarefa"
   ClientHeight    =   6960
   ClientLeft      =   2595
   ClientTop       =   2760
   ClientWidth     =   6855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin iGrid251_75B4A91C.iGrid GrdVenda 
      Height          =   2655
      Left            =   7440
      TabIndex        =   25
      Top             =   2760
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4683
      BorderStyle     =   1
      HighlightBackColorNoFocus=   14737632
   End
   Begin XtremeSuiteControls.DateTimePicker CmbDTTAREFA 
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   1575
      _Version        =   720898
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   68
      Format          =   1
      CurrentDate     =   40356.1843055556
   End
   Begin XtremeSuiteControls.PushButton CmdCancel 
      Height          =   375
      Left            =   4080
      TabIndex        =   20
      Top             =   6480
      Width           =   1215
      _Version        =   720898
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Cancelar"
      ForeColor       =   192
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdOk 
      Height          =   375
      Left            =   5520
      TabIndex        =   19
      Top             =   6480
      Width           =   1215
      _Version        =   720898
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Ok"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GrpSessao 
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   960
      Width           =   6495
      _Version        =   720898
      _ExtentX        =   11456
      _ExtentY        =   1296
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit TxtTEL1 
         Height          =   315
         Left            =   5280
         TabIndex        =   6
         Top             =   0
         Width           =   1215
         _Version        =   720898
         _ExtentX        =   2143
         _ExtentY        =   564
         _StockProps     =   77
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtNOME 
         Height          =   315
         Left            =   720
         TabIndex        =   3
         Top             =   0
         Width           =   2895
         _Version        =   720898
         _ExtentX        =   5106
         _ExtentY        =   564
         _StockProps     =   77
         BackColor       =   16777215
         Enabled         =   0   'False
         Text            =   "Patricia Moreira"
         BackColor       =   16777215
         Appearance      =   1
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdIDCLIENTE 
         Height          =   315
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   705
         _Version        =   720898
         _ExtentX        =   1244
         _ExtentY        =   564
         _StockProps     =   79
         Caption         =   "Cliente"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   10
         ImageAlignment  =   6
         TextImageRelation=   0
      End
      Begin XtremeSuiteControls.PushButton CmdLovCli 
         Height          =   320
         Left            =   3600
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
         _Version        =   720898
         _ExtentX        =   661
         _ExtentY        =   564
         _StockProps     =   79
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
         Picture         =   "FrmCadTarefa.frx":0000
      End
      Begin XtremeSuiteControls.PushButton CmdEMAIL 
         Height          =   255
         Left            =   720
         TabIndex        =   26
         Top             =   420
         Width           =   5655
         _Version        =   720898
         _ExtentX        =   9975
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "aaaaaa@aaa.com.br"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   0
         Appearance      =   2
         MultiLine       =   0   'False
         ImageAlignment  =   1
         BorderGap       =   0
         ImageGap        =   0
      End
      Begin XtremeSuiteControls.Label LblMail 
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   465
         _Version        =   720898
         _ExtentX        =   820
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "e-Mail:"
         ForeColor       =   8421504
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin VB.Label LblTel2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Celular/Tel."
         Height          =   195
         Left            =   4440
         TabIndex        =   5
         Top             =   0
         Width           =   825
      End
   End
   Begin XtremeSuiteControls.PushButton CmdExcluir 
      Height          =   375
      Left            =   120
      TabIndex        =   21
      ToolTipText     =   "Excluir"
      Top             =   6480
      Width           =   1215
      _Version        =   720898
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Excluir"
      ForeColor       =   64
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
      Picture         =   "FrmCadTarefa.frx":0183
   End
   Begin XtremeSuiteControls.FlatEdit TxtTITULO 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6495
      _Version        =   720898
      _ExtentX        =   11456
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   16777215
      Text            =   "Patricia Moreira"
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox CmbTPTAREFA 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   4155
      _Version        =   720898
      _ExtentX        =   7329
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtIDTAREFA 
      Height          =   315
      Left            =   6120
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   735
      _Version        =   720898
      _ExtentX        =   1296
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   16777215
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtDSCTAREFA 
      Height          =   2625
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   6615
      _Version        =   720898
      _ExtentX        =   11668
      _ExtentY        =   4630
      _StockProps     =   77
      BackColor       =   16777215
      Text            =   "Patricia Moreira"
      BackColor       =   16777215
      MultiLine       =   -1  'True
      ScrollBars      =   2
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox CmbIDEVENTO 
      Height          =   315
      Left            =   4560
      TabIndex        =   10
      Top             =   2160
      Width           =   2115
      _Version        =   720898
      _ExtentX        =   3731
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtIDEVENTO 
      Height          =   195
      Left            =   5640
      TabIndex        =   23
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
      _Version        =   720898
      _ExtentX        =   1296
      _ExtentY        =   344
      _StockProps     =   77
      BackColor       =   16777215
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox CmbSITTAREFA 
      Height          =   315
      Left            =   4560
      TabIndex        =   16
      Top             =   3000
      Width           =   2115
      _Version        =   720898
      _ExtentX        =   3731
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdEvento 
      Height          =   285
      Left            =   4560
      TabIndex        =   9
      Top             =   1920
      Width           =   705
      _Version        =   720898
      _ExtentX        =   1244
      _ExtentY        =   494
      _StockProps     =   79
      Caption         =   "Evento"
      Transparent     =   -1  'True
      TextAlignment   =   10
      Appearance      =   4
      TextImageRelation=   0
   End
   Begin XtremeSuiteControls.ComboBox CmbPrioridade 
      Height          =   315
      Left            =   2640
      TabIndex        =   14
      Top             =   3000
      Width           =   1275
      _Version        =   720898
      _ExtentX        =   2249
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdeMail2 
      Height          =   375
      Left            =   1800
      TabIndex        =   28
      ToolTipText     =   "Excluir"
      Top             =   6480
      Width           =   1815
      _Version        =   720898
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "e-Mail Padrão"
      ForeColor       =   64
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
      Picture         =   "FrmCadTarefa.frx":0C4D
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   120
      X2              =   6720
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label LblSIT 
      AutoSize        =   -1  'True
      Caption         =   "Situação"
      Height          =   195
      Left            =   4560
      TabIndex        =   15
      Top             =   2760
      Width           =   630
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   120
      X2              =   6720
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   120
      X2              =   6720
      Y1              =   1840
      Y2              =   1840
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Título"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   420
   End
   Begin VB.Label LblPrioridade 
      AutoSize        =   -1  'True
      Caption         =   "Prioridade"
      Height          =   195
      Left            =   2640
      TabIndex        =   13
      Top             =   2760
      Width           =   705
   End
   Begin VB.Label LblPagamento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detalhes"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Width           =   630
   End
   Begin VB.Label LblDTATEND 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   345
   End
End
Attribute VB_Name = "FrmCadTarefa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://www.guaru.com.br/sistemas/document/pdvtef_06.asp
Option Explicit
Event Activate()
Event Load()
Event KeyUp(KeyCode As Integer, Shift As Integer)
'Event Unload(Cancel As Integer)
'Event Resize()
'
Event CmbTPTAREFAClick()
Event CmbTPTAREFALostFocus()
Event CmdOkClick()
Event CmdCancelClick()
Event CmdExcluirClick()
Event CmdEventoClick()
Event CmdEMAILClick()
Event CmdEMAIL2Click()
Event CmdIDCLIENTEClick()
Event CmdLovCliClick()
Event TxtDSCTAREFAChange()
Event TxtDSCTAREFAKeyPress(KeyAscii As Integer)
Event TxtNOMEChange()
Event TxtTEL1Change()
Private Sub CmbTPTAREFA_Click()
   RaiseEvent CmbTPTAREFAClick
End Sub
Private Sub CmbTPTAREFA_LostFocus()
   RaiseEvent CmbTPTAREFALostFocus
End Sub
Private Sub cmdCancel_Click()
   RaiseEvent CmdCancelClick
End Sub
Private Sub CmdEMAIL_Click()
   RaiseEvent CmdEMAILClick
End Sub

Private Sub CmdeMail2_Click()
   RaiseEvent CmdEMAIL2Click
End Sub
Private Sub CmdEvento_Click()
   RaiseEvent CmdEventoClick
End Sub
Private Sub CmdExcluir_Click()
   RaiseEvent CmdExcluirClick
End Sub
Private Sub CmdIDCLIENTE_Click()
   RaiseEvent CmdIDCLIENTEClick
End Sub
Private Sub CmdLovCli_Click()
   RaiseEvent CmdLovCliClick
End Sub
Private Sub cmdOk_Click()
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
Private Sub TxtDSCTAREFA_Change()
   RaiseEvent TxtDSCTAREFAChange
End Sub
Private Sub TxtDSCTAREFA_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtDSCTAREFAKeyPress(KeyAscii)
End Sub
Private Sub TxtNOME_Change()
   RaiseEvent TxtNOMEChange
End Sub
Private Sub TxtNOME_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtTEL1_Change()
   RaiseEvent TxtTEL1Change
End Sub
Private Sub TxtTEL1_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub

