VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#11.2#0"; "Codejock.ReportControl.v11.2.2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#11.2#0"; "Codejock.CommandBars.v11.2.2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmContasRP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Contas a Receber"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeReportControl.ReportControl GrdContasRP 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6495
      _Version        =   720898
      _ExtentX        =   11456
      _ExtentY        =   8916
      _StockProps     =   64
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContasRP.frx":0000
            Key             =   "K1"
            Object.Tag             =   "01"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltrar 
      Height          =   330
      Left            =   4080
      TabIndex        =   1
      Top             =   30
      Width           =   2130
      _Version        =   720898
      _ExtentX        =   3757
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   12632256
      BackColor       =   -2147483643
      Text            =   "Pesquisar Vendas"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Image imgLupa 
      Height          =   270
      Left            =   6240
      Picture         =   "FrmContasRP.frx":0ACA
      Stretch         =   -1  'True
      Top             =   60
      Width           =   255
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   1320
      Top             =   5760
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager ImgToobar 
      Left            =   3600
      Top             =   5760
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "FrmContasRP.frx":0F1C
   End
End
Attribute VB_Name = "FrmContasRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Load()
Event Resize()
Event Unload(Cancel As Integer)
Event CommandBarsExecute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Event txtFiltrarGotFocus()
Event txtFiltrarLostFocus()
Event txtFiltrarKeyPress(KeyAscii As Integer)
Event GrdContasRPBeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
Event GrdContasRPKeyUp(KeyCode As Integer, Shift As Integer)
Event GrdContasRPRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Event GrdContasRPSelectionChanged()
Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   RaiseEvent CommandBarsExecute(Control)
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
Private Sub Form_Resize()
   RaiseEvent Resize
End Sub
Private Sub Form_Unload(Cancel As Integer)
   RaiseEvent Unload(Cancel)
End Sub
Private Sub GrdContasRP_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
   RaiseEvent GrdContasRPBeforeDrawRow(Row, Item, Metrics)
End Sub
Private Sub GrdContasRP_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent GrdContasRPKeyUp(KeyCode, Shift)
End Sub
Private Sub GrdContasRP_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   RaiseEvent GrdContasRPRowDblClick(Row, Item)
End Sub
Private Sub GrdContasRP_SelectionChanged()
  RaiseEvent GrdContasRPSelectionChanged
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


