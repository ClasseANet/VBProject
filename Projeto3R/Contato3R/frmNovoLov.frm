VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#11.2#0"; "Codejock.ReportControl.v11.2.2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Begin VB.Form frmNovoLov 
   Caption         =   "Listagem para Seleção"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin XtremeReportControl.ReportControl grdItens 
      Height          =   4155
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   7080
      _Version        =   720898
      _ExtentX        =   12488
      _ExtentY        =   7329
      _StockProps     =   64
   End
   Begin XtremeSuiteControls.CheckBox ChkEmEspera 
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   4320
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Em Espera"
      BackColor       =   -2147483643
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   5
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox ChkInativo 
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   4320
      Width           =   855
      _Version        =   720898
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Inativos"
      BackColor       =   -2147483643
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   5
   End
   Begin XtremeSuiteControls.PushButton cmdFiltrar 
      Height          =   375
      Left            =   2925
      TabIndex        =   5
      Top             =   4815
      Width           =   375
      _Version        =   720898
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Picture         =   "frmNovoLov.frx":0000
   End
   Begin XtremeSuiteControls.PushButton cmdSelecionar 
      Height          =   420
      Left            =   4410
      TabIndex        =   2
      Top             =   4800
      Width           =   1320
      _Version        =   720898
      _ExtentX        =   2328
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "S&elecionar"
      ForeColor       =   0
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdSair 
      Height          =   420
      Left            =   5895
      TabIndex        =   3
      Top             =   4800
      Width           =   1320
      _Version        =   720898
      _ExtentX        =   2328
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "&Sair"
      ForeColor       =   16711680
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox ChkAtivo 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   735
      _Version        =   720898
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Ativos"
      BackColor       =   -2147483643
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   5
      Value           =   1
   End
   Begin XtremeSuiteControls.FlatEdit TxtFiltrar 
      Height          =   315
      Left            =   600
      TabIndex        =   4
      Top             =   4840
      Width           =   2280
      _Version        =   720898
      _ExtentX        =   4022
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   16777215
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblFiltrar 
      Height          =   330
      Left            =   135
      TabIndex        =   1
      Top             =   4815
      Width           =   510
      _Version        =   720898
      _ExtentX        =   900
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Filtar:"
   End
End
Attribute VB_Name = "frmNovoLov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event Resize()
Event FormKeyUp(KeyCode As Integer, Shift As Integer)
Event FormKeyPress(KeyAscii As Integer)
Event CmdSairClick()
Event CmdSelecionarClick()
Event CmdFiltrarClick()
Event ChkEmEsperaClick()
Event ChkAtivoClick()
Event ChkInativoClick()
Event grdItensKeyDown(KeyCode As Integer, Shift As Integer)
Event grdItensKeyUp(KeyCode As Integer, Shift As Integer)
Event TxtFiltrarKeyDown(KeyCode As Integer, Shift As Integer)
Event TxtFiltrarKeyUp(KeyCode As Integer, Shift As Integer)
Event TxtFiltrarKeyPress(KeyAscii As Integer)
Private Sub ChkAtivo_Click()
   RaiseEvent ChkAtivoClick
End Sub
Private Sub ChkEmEspera_Click()
   RaiseEvent ChkEmEsperaClick
End Sub
Private Sub ChkInativo_Click()
   RaiseEvent ChkInativoClick
End Sub
Private Sub CmdSair_Click()
   RaiseEvent CmdSairClick
End Sub
Private Sub cmdSelecionar_Click()
   RaiseEvent CmdSelecionarClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   RaiseEvent FormKeyPress(KeyAscii)
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent FormKeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Form_Resize()
   RaiseEvent Resize
End Sub
Private Sub CmdFiltrar_Click()
   RaiseEvent CmdFiltrarClick
End Sub
Private Sub grdItens_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent grdItensKeyDown(KeyCode, Shift)
End Sub
Private Sub grdItens_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent grdItensKeyUp(KeyCode, Shift)
End Sub
Private Sub grdItens_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   RaiseEvent CmdSelecionarClick
End Sub
Private Sub txtFiltrar_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent TxtFiltrarKeyDown(KeyCode, Shift)
End Sub
Private Sub txtFiltrar_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtFiltrarKeyPress(KeyAscii)
End Sub
Private Sub txtFiltrar_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent TxtFiltrarKeyUp(KeyCode, Shift)
End Sub
