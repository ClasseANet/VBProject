VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.ShortcutBar.v11.2.2.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#11.2#0"; "Codejock.ReportControl.v11.2.2.ocx"
Begin VB.Form FrmCadastro 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Cadastro Geral"
   ClientHeight    =   7605
   ClientLeft      =   18165
   ClientTop       =   1695
   ClientWidth     =   6585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin XtremeReportControl.ReportControl GrdCadastro 
      Height          =   5055
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   6495
      _Version        =   720898
      _ExtentX        =   11456
      _ExtentY        =   8916
      _StockProps     =   64
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltrar 
      Height          =   330
      Left            =   3915
      TabIndex        =   0
      Top             =   60
      Width           =   2130
      _Version        =   720898
      _ExtentX        =   3757
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   12632256
      BackColor       =   -2147483643
      Text            =   "Pesquisar..."
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption SccContato 
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6720
      _Version        =   720898
      _ExtentX        =   11853
      _ExtentY        =   741
      _StockProps     =   14
      Caption         =   "Titulo"
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
   Begin VB.Image imgLupa 
      Height          =   270
      Left            =   6075
      Picture         =   "FrmCADASTRO.frx":0000
      Stretch         =   -1  'True
      Top             =   60
      Width           =   255
   End
End
Attribute VB_Name = "FrmCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Load()
Event Rezise()
Event Unload(Cancel As Integer)
Event txtFiltrarGotFocus()
Event txtFiltrarLostFocus()
Event txtFiltrarKeyPress(KeyAscii As Integer)
Event GrdCadastroRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Form_Resize()
   RaiseEvent Rezise
End Sub
Private Sub Form_Unload(Cancel As Integer)
   RaiseEvent Unload(Cancel)
End Sub
Private Sub GrdCadastro_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub GrdCadastro_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   RaiseEvent GrdCadastroRowDblClick(Row, Item)
End Sub
Private Sub txtFiltrar_GotFocus()
   RaiseEvent txtFiltrarGotFocus
End Sub
Private Sub txtFiltrar_KeyPress(KeyAscii As Integer)
   RaiseEvent txtFiltrarKeyPress(KeyAscii)
End Sub
Private Sub txtFiltrar_LostFocus()
   RaiseEvent txtFiltrarLostFocus
End Sub
