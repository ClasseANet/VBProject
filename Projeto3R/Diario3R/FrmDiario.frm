VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.ShortcutBar.v11.2.2.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#11.2#0"; "Codejock.ReportControl.v11.2.2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#11.2#0"; "Codejock.CommandBars.v11.2.2.ocx"
Begin VB.Form FrmDiario 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Conta"
   ClientHeight    =   6225
   ClientLeft      =   3645
   ClientTop       =   840
   ClientWidth     =   9285
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin XtremeReportControl.ReportControl GrdDiario 
      Height          =   2535
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   3375
      _Version        =   720898
      _ExtentX        =   5953
      _ExtentY        =   4471
      _StockProps     =   64
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4320
      Top             =   3480
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltrar 
      Height          =   330
      Left            =   6720
      TabIndex        =   1
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
   Begin XtremeSuiteControls.FlatEdit TxtDSCDIARIO 
      Height          =   2055
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   4140
      _Version        =   720898
      _ExtentX        =   7302
      _ExtentY        =   3625
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "TEste de fonte"
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeCommandBars.ImageManager ImgShortcutBar 
      Left            =   4440
      Top             =   120
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "FrmDiario.frx":0000
   End
   Begin VB.Image imgLupa 
      Height          =   270
      Left            =   8880
      Picture         =   "FrmDiario.frx":18A2
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
      Caption         =   "Ocorrência"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.76
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmDiario"
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
Event txtFiltrarGotFocus()
Event txtFiltrarLostFocus()
Event txtFiltrarKeyPress(KeyAscii As Integer)

Event GrdDiariosRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Event GrdDiariosBeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
Event GrdDiariosSelectionChanged()
Event GrdDiariosMouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
   
Event TxtDSCDIARIOLostFocus()
Event TxtDSCDIARIOChange()
Event TxtDSCDIARIOKeyUp(KeyCode As Integer, Shift As Integer)
Event Timer1()
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub

Private Sub Form_Deactivate()
' RaiseEvent TxtDSCDIARIOLostFocus
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
Private Sub GrdDiario_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
   RaiseEvent GrdDiariosBeforeDrawRow(Row, Item, Metrics)
End Sub
Private Sub GrdDiario_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub GrdDiario_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
   RaiseEvent GrdDiariosMouseDown(Button, Shift, x, y)
End Sub
Private Sub GrdDiario_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   RaiseEvent GrdDiariosRowDblClick(Row, Item)
End Sub
Private Sub GrdDiario_SelectionChanged()
   RaiseEvent GrdDiariosSelectionChanged
End Sub

Private Sub Timer1_Timer()
   RaiseEvent Timer1
End Sub

Private Sub TxtDSCDIARIO_Change()
   RaiseEvent TxtDSCDIARIOChange
End Sub
Private Sub TxtDSCDIARIO_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent TxtDSCDIARIOKeyUp(KeyCode, Shift)
End Sub
Private Sub TxtDSCDIARIO_LostFocus()
   RaiseEvent TxtDSCDIARIOLostFocus
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
