VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.ShortcutBar.v11.2.2.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#11.2#0"; "Codejock.ReportControl.v11.2.2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmContato 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Contatos"
   ClientHeight    =   6060
   ClientLeft      =   6030
   ClientTop       =   2070
   ClientWidth     =   6810
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin XtremeReportControl.ReportControl GrdContato 
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6495
      _Version        =   720898
      _ExtentX        =   11456
      _ExtentY        =   8916
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContato.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContato.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContato.frx":0734
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContato.frx":0ACE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltrar 
      Height          =   330
      Left            =   3915
      TabIndex        =   2
      Top             =   60
      Width           =   2130
      _Version        =   720898
      _ExtentX        =   3757
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   12632256
      BackColor       =   -2147483643
      Text            =   "Pesquisar Contatos"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label LblStGrd 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total: (0 Itens)"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   6495
   End
   Begin VB.Image imgLupa 
      Height          =   270
      Left            =   6075
      Picture         =   "FrmContato.frx":0C28
      Stretch         =   -1  'True
      Top             =   60
      Width           =   255
   End
   Begin XtremeShortcutBar.ShortcutCaption SccContato 
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6720
      _Version        =   720898
      _ExtentX        =   11853
      _ExtentY        =   741
      _StockProps     =   14
      Caption         =   "Contatos"
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
Attribute VB_Name = "FrmContato"
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
Event TxtFiltrarKeyPress(KeyAscii As Integer)
Event GrdContatoRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
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
Private Sub GRdContato_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub GrdContato_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   RaiseEvent GrdContatoRowDblClick(Row, Item)
End Sub

Private Sub GrdContato_SortOrderChanged()
'   Dim i As Integer
'   For i = 0 To Me.GrdContato.Records.Count - 1
'      'Me.GrdContato.Records(i).Item(0).Value = i + 1
'      Me.GrdContato.Rows(i).Record.Item(0).Value = i + 1
'   Next
'   Me.GrdContato.Rows.Row(0).Selected = True
'   Me.GrdContato.Populate
End Sub

Private Sub txtFiltrar_GotFocus()
   RaiseEvent txtFiltrarGotFocus
End Sub
Private Sub txtFiltrar_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtFiltrarKeyPress(KeyAscii)
End Sub
Private Sub txtFiltrar_LostFocus()
   RaiseEvent txtFiltrarLostFocus
End Sub

