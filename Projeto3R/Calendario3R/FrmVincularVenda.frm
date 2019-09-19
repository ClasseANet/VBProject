VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Begin VB.Form FrmVincularVenda 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Vincular Venda <-> Atendimento"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView LstVendas 
      Height          =   3615
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   8415
      _Version        =   720898
      _ExtentX        =   14843
      _ExtentY        =   6376
      _StockProps     =   77
      BackColor       =   -2147483643
      Checkboxes      =   -1  'True
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GrpFiltros 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _Version        =   720898
      _ExtentX        =   14843
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   " FILTROS "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox ChkCliente 
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   3855
         _Version        =   720898
         _ExtentX        =   6800
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Apenas vendas do cliente"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox ChkVendas 
         Height          =   375
         Left            =   4560
         TabIndex        =   2
         Top             =   240
         Width           =   3735
         _Version        =   720898
         _ExtentX        =   6588
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Serviços do Atendimento"
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
         Value           =   1
      End
   End
   Begin XtremeSuiteControls.PushButton CmdCancel 
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   8640
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Cancelar"
      ForeColor       =   192
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdOk 
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   8640
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Ok"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
   End
   Begin iGrid251_75B4A91C.iGrid GrdItens 
      Height          =   3615
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6376
      Appearance      =   0
      BorderStyle     =   1
      HighlightBackColorNoFocus=   14737632
      KeyPressBehaviour=   1
   End
   Begin XtremeSuiteControls.PushButton CmdReconstruir 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   8640
      Visible         =   0   'False
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Reconstruir"
      ForeColor       =   16384
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   1695
      _Version        =   720898
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "ATENDIMENTO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label LblVendas 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1695
      _Version        =   720898
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "VENDAS"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmVincularVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event Load()
Event Unload(Cancel As Integer)
Event Resize()
Event CmdOkClick()
Event CmdCancelClick()
Event CmdReconstruirClick()
Event ChkClienteClick()
Event ChkVendasClick()
Event GrdItensAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
Event GrdItensCancelEdit(ByVal lRow As Long, ByVal lCol As Long)
Event GrdItensBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
Event GrdItensKeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
Event GrdItensKeyPress(KeyAscii As Integer)
Event GrdItensKeyUp(KeyCode As Integer, Shift As Integer)
Event LblVendasDblClick()
Event LstVendasItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Private Sub ChkCliente_Click()
   RaiseEvent ChkClienteClick
End Sub
Private Sub ChkVendas_Click()
   RaiseEvent ChkVendasClick
End Sub
Private Sub cmdCancel_Click()
   RaiseEvent CmdCancelClick
End Sub
Private Sub cmdOk_Click()
   RaiseEvent CmdOkClick
End Sub

Private Sub CmdReconstruir_Click()
   RaiseEvent CmdReconstruirClick
End Sub

Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
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
Private Sub GrdItens_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdItensAfterCommitEdit(lRow, lCol)
End Sub
Private Sub GrdItens_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   RaiseEvent GrdItensBeforeCommitEdit(lRow, lCol, eResult, sNewText, vNewValue, lConvErr)
End Sub
Private Sub GrdItens_CancelEdit(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdItensCancelEdit(lRow, lCol)
End Sub
Private Sub GrdItens_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
   RaiseEvent GrdItensKeyDown(KeyCode, Shift, bDoDefault)
End Sub
Private Sub GrdItens_KeyPress(KeyAscii As Integer)
   RaiseEvent GrdItensKeyPress(KeyAscii)
End Sub
Private Sub GrdItens_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent GrdItensKeyUp(KeyCode, Shift)
End Sub
Private Sub LblVendas_DblClick()
   RaiseEvent LblVendasDblClick
End Sub

Private Sub LstVendas_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
   If InArray(ColumnHeader.Index, Array(1, 4)) Then
      If Me.LstVendas.SortKey = ColumnHeader.Index - 1 Then
         Me.LstVendas.SortKey = ColumnHeader.Index - 1
         Me.LstVendas.SortOrder = IIf(Me.LstVendas.SortOrder = 1, 2, 1)
      Else
         Me.LstVendas.SortKey = ColumnHeader.Index - 1
         Me.LstVendas.SortOrder = 2
      End If
   End If
End Sub
Private Sub LstVendas_DblClick()
   If Not Me.LstVendas.SelectedItem Is Nothing Then
      Me.LstVendas.SelectedItem.Checked = Not Me.LstVendas.SelectedItem.Checked
   End If
End Sub

Private Sub LstVendas_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
   RaiseEvent LstVendasItemCheck(Item)
End Sub

Private Sub PushButton1_Click()

End Sub
