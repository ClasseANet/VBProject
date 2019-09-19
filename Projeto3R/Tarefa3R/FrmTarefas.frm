VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.ShortcutBar.v11.2.2.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#11.2#0"; "Codejock.ReportControl.v11.2.2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#11.2#0"; "Codejock.CommandBars.v11.2.2.ocx"
Begin VB.Form FrmTarefas 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Conta"
   ClientHeight    =   7215
   ClientLeft      =   3645
   ClientTop       =   840
   ClientWidth     =   10200
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin XtremeReportControl.ReportControl GrdTarefas 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   9975
      _Version        =   720898
      _ExtentX        =   17595
      _ExtentY        =   8070
      _StockProps     =   64
   End
   Begin XtremeSuiteControls.GroupBox GrpEVENTO 
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   9975
      _Version        =   720898
      _ExtentX        =   17595
      _ExtentY        =   3201
      _StockProps     =   79
      ForeColor       =   -2147483636
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   1
      BorderStyle     =   2
      Begin XtremeSuiteControls.GroupBox GrpTarefa 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   9735
         _Version        =   720898
         _ExtentX        =   17171
         _ExtentY        =   661
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   4
         Begin XtremeSuiteControls.Label LblDSCTPTAREFA 
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   75
            Width           =   9480
            _Version        =   720898
            _ExtentX        =   16722
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "CONFIRMAR AGENDA"
            ForeColor       =   8421504
            BackColor       =   16777215
            Alignment       =   2
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton CmdEMAIL 
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   1560
         Width           =   1950
         _Version        =   720898
         _ExtentX        =   3440
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "aaaaaa@aaa.com.br"
         ForeColor       =   16711680
         BackColor       =   16777215
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
         Appearance      =   2
         MultiLine       =   0   'False
         ImageAlignment  =   4
         BorderGap       =   0
         ImageGap        =   0
      End
      Begin XtremeSuiteControls.FlatEdit TxtOutros 
         Height          =   735
         Left            =   6480
         TabIndex        =   20
         Top             =   795
         Width           =   3135
         _Version        =   720898
         _ExtentX        =   5530
         _ExtentY        =   1296
         _StockProps     =   77
         ForeColor       =   8421504
         BackColor       =   -2147483643
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   4
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblSTARTTIME 
         Height          =   240
         Left            =   7920
         TabIndex        =   19
         Top             =   435
         Width           =   1695
         _Version        =   720898
         _ExtentX        =   2990
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "88/88/8888 88:88h"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblEvento 
         Height          =   195
         Left            =   7320
         TabIndex        =   18
         Top             =   480
         Width           =   555
         _Version        =   720898
         _ExtentX        =   979
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Evento:"
         ForeColor       =   8421504
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblEMAIL 
         Height          =   240
         Left            =   1080
         TabIndex        =   16
         Top             =   1560
         Width           =   1950
         _Version        =   720898
         _ExtentX        =   3440
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "aaaaaa@aaa.com.br"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblMail 
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1560
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
      Begin XtremeSuiteControls.Label LblFAX 
         Height          =   240
         Left            =   1080
         TabIndex        =   14
         Top             =   1200
         Width           =   900
         _Version        =   720898
         _ExtentX        =   1588
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "1234-5678"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblOutroTel 
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   750
         _Version        =   720898
         _ExtentX        =   1323
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Outro Tel.:"
         ForeColor       =   8421504
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblTEL2 
         Height          =   240
         Left            =   3360
         TabIndex        =   12
         Top             =   840
         Width           =   900
         _Version        =   720898
         _ExtentX        =   1588
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "1234-5678"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   195
         Left            =   2520
         TabIndex        =   11
         Top             =   840
         Width           =   675
         _Version        =   720898
         _ExtentX        =   1191
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Telefone:"
         ForeColor       =   8421504
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblTEL1 
         Height          =   240
         Left            =   1080
         TabIndex        =   10
         Top             =   840
         Width           =   900
         _Version        =   720898
         _ExtentX        =   1588
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "1234-5678"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblCelular 
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   525
         _Version        =   720898
         _ExtentX        =   926
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Celular:"
         ForeColor       =   8421504
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblNMCLIENTE 
         Height          =   240
         Left            =   1080
         TabIndex        =   8
         Top             =   480
         Width           =   1470
         _Version        =   720898
         _ExtentX        =   2593
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Nome do cliente"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblCLiente 
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   525
         _Version        =   720898
         _ExtentX        =   926
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Cliente:"
         ForeColor       =   8421504
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblSituacao 
         Height          =   405
         Left            =   6960
         TabIndex        =   4
         Top             =   1460
         Width           =   2760
         _Version        =   720898
         _ExtentX        =   4868
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Cancelado"
         ForeColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltrar 
      Height          =   330
      Left            =   6720
      TabIndex        =   3
      Top             =   60
      Width           =   2130
      _Version        =   720898
      _ExtentX        =   3757
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   12632256
      BackColor       =   -2147483643
      Text            =   "Pesquisar Extrato"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeCommandBars.ImageManager ImgShortcutBar 
      Left            =   4440
      Top             =   120
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "FrmTarefas.frx":0000
   End
   Begin VB.Image ImgBandeira 
      Height          =   480
      Left            =   0
      Picture         =   "FrmTarefas.frx":18A2
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgLupa 
      Height          =   270
      Left            =   8880
      Picture         =   "FrmTarefas.frx":216C
      Stretch         =   -1  'True
      Top             =   60
      Width           =   255
   End
   Begin XtremeShortcutBar.ShortcutCaption SccConta 
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10200
      _Version        =   720898
      _ExtentX        =   17992
      _ExtentY        =   741
      _StockProps     =   14
      Caption         =   "     Lista de Tarefas"
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
End
Attribute VB_Name = "FrmTarefas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event Load()
Event Rezise()
Event Unload(Cancel As Integer)
'Event CmdCancelarClick()
'Event CmdSalvarClick()
'Event CmdNovoClick()
Event CmdNumClick()
Event CmdEMAILClick()
'Event CmbDeLostFocus()
'Event CmbParaLostFocus()
'Event CmbFavorecidoLostFocus()
'Event CmbCategoriaLostFocus()
'Event CmbCategoriaChange()
'Event CmbCategoriaClick()
'Event CmbSubCategoriaLostFocus()
'Event TxtDTBAIXALostFocus()
'Event CmbDTBAIXALostFocus()
'Event CmbDTBAIXAChange()
Event LblEMAILClick()
Event txtFiltrarGotFocus()
Event txtFiltrarLostFocus()
Event txtFiltrarKeyPress(KeyAscii As Integer)
'Event TxtNDOCLostFocus()
Event GrdTarefasRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Event GrdTarefasBeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
Event GrdTarefasSelectionChanged()
Event GrdTarefasMouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
Private Sub CmdEMAIL_Click()
   RaiseEvent CmdEMAILClick
End Sub

Private Sub CmdNum_Click()
   RaiseEvent CmdNumClick
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
Private Sub GrdTarefas_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
   RaiseEvent GrdTarefasBeforeDrawRow(Row, Item, Metrics)
End Sub
Private Sub GrdTarefas_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub GrdTarefas_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
   RaiseEvent GrdTarefasMouseDown(Button, Shift, x, y)
End Sub
Private Sub GrdTarefas_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   RaiseEvent GrdTarefasRowDblClick(Row, Item)
End Sub
Private Sub GrdTarefas_SelectionChanged()
   RaiseEvent GrdTarefasSelectionChanged
End Sub
Private Sub LblEMAIL_Click()
   RaiseEvent LblEMAILClick
End Sub

Private Sub txtFiltrar_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
   RaiseEvent txtFiltrarGotFocus
End Sub
Private Sub txtFiltrar_KeyPress(KeyAscii As Integer)
   RaiseEvent txtFiltrarKeyPress(KeyAscii)
End Sub
Private Sub txtFiltrar_LostFocus()
   RaiseEvent txtFiltrarLostFocus
End Sub

