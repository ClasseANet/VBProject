VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.ShortcutBar.v11.2.2.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#11.2#0"; "Codejock.ReportControl.v11.2.2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Begin VB.Form FrmProdutos 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Lista de Produtos"
   ClientHeight    =   5055
   ClientLeft      =   15720
   ClientTop       =   1770
   ClientWidth     =   8070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin XtremeReportControl.ReportControl GrdProd 
      Height          =   2775
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   5655
      _Version        =   720898
      _ExtentX        =   9975
      _ExtentY        =   4895
      _StockProps     =   64
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltrar 
      Height          =   330
      Left            =   5520
      TabIndex        =   2
      Top             =   40
      Width           =   2130
      _Version        =   720898
      _ExtentX        =   3757
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   12632256
      BackColor       =   -2147483643
      Text            =   "Pesquisar Produtos"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.PictureBox PictBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   480
      ScaleHeight     =   915
      ScaleWidth      =   6870
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   6870
      Begin XtremeSuiteControls.GroupBox GrpBoxBottom 
         Height          =   975
         Left            =   -600
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   8175
         _Version        =   720898
         _ExtentX        =   14420
         _ExtentY        =   1720
         _StockProps     =   79
         Transparent     =   -1  'True
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton CmdSair 
            Height          =   375
            Left            =   5880
            TabIndex        =   5
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
            _Version        =   720898
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Sai&r"
            ForeColor       =   0
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdExcluir 
            Height          =   375
            Left            =   2640
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "&Excluir"
            ForeColor       =   192
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdNovo 
            Height          =   375
            Left            =   840
            TabIndex        =   7
            Top             =   240
            Visible         =   0   'False
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "&Novo"
            ForeColor       =   12582912
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdEditar 
            Height          =   375
            Left            =   4200
            TabIndex        =   8
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
            _Version        =   720898
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "&Editar"
            ForeColor       =   4210752
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.TabControlPage TabPgBotton 
            Height          =   855
            Left            =   -1080
            TabIndex        =   9
            Top             =   120
            Visible         =   0   'False
            Width           =   8895
            _Version        =   720898
            _ExtentX        =   15690
            _ExtentY        =   1508
            _StockProps     =   1
         End
      End
   End
   Begin VB.Image imgLupa 
      Height          =   270
      Left            =   7680
      Picture         =   "FrmProdutos.frx":0000
      Stretch         =   -1  'True
      Top             =   40
      Width           =   255
   End
   Begin XtremeShortcutBar.ShortcutCaption SccTit 
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10200
      _Version        =   720898
      _ExtentX        =   17992
      _ExtentY        =   741
      _StockProps     =   14
      Caption         =   "     Produtos"
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
Attribute VB_Name = "FrmProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event GrdProdKeyUp(KeyCode, Shift)
Event GrdProdRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Event Load()
Event Rezise()
Event Unload(Cancel As Integer)

Event CmdExcluirClick()
Event CmdEditarClick()
Event CmdNovoClick()
Private Sub CmdExcluir_Click()
   RaiseEvent CmdExcluirClick
End Sub
Private Sub CmdNovo_Click()
   RaiseEvent CmdNovoClick
End Sub
Private Sub CmdEditar_Click()
   RaiseEvent CmdEditarClick
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
Private Sub GrdProd_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent GrdProdKeyUp(KeyCode, Shift)
End Sub
Private Sub GrdProd_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   RaiseEvent GrdProdRowDblClick(Row, Item)
End Sub

