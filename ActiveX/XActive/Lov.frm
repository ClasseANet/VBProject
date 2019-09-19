VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Begin VB.Form FrmLov 
   AutoRedraw      =   -1  'True
   Caption         =   "GrdLov"
   ClientHeight    =   6780
   ClientLeft      =   2760
   ClientTop       =   510
   ClientWidth     =   5670
   Icon            =   "Lov.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6780
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin Crystal.CrystalReport CryRprt 
      Left            =   4560
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox TxtLov 
      Height          =   330
      Left            =   1140
      TabIndex        =   2
      Top             =   5880
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid GrdLov 
      Height          =   5355
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   9446
      _Version        =   393216
      BackColor       =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSComctlLib.TreeView TreLOV 
      Height          =   5355
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9446
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin XtremeSuiteControls.PushButton CmdFiltrar 
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   5880
      Width           =   855
      _Version        =   720898
      _ExtentX        =   1508
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Filtrar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdToExcel 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   615
      _Version        =   720898
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Excel"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdTreeView 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   0
      Width           =   1215
      _Version        =   720898
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Estrutura"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdGrid 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   0
      Width           =   1215
      _Version        =   720898
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Tabela"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdOk 
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   6360
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Ok"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdCancel 
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   6360
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cancelar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdImprimir 
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   0
      Width           =   735
      _Version        =   720898
      _ExtentX        =   1296
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Imprimir"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label LblLov 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Localizar :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   3
      Left            =   60
      TabIndex        =   3
      ToolTipText     =   "* Use [Esc] para limpar o campo ""Localizar"""
      Top             =   5880
      Width           =   1050
   End
End
Attribute VB_Name = "FrmLov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event Activate()
Event Load()
Event Resize()
Event QueryUnload(Cancel, UnloadMode)
Event Unload(Cancel As Integer)
Event CmdOkClick()
Event CmdCancelClick()
Event CmdGridClick()
Event CmdTreeViewClick()
Event CmdFiltrarClick()
Event CmdImprimirClick()
Event CmdToExcelClick()
Event Excluir()
Event GrdClick()
Event GrdCompare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
Event GrdDblClick()
Event GrdLovSelChange()
Event GrdLovLeaveCell()
Event FrmKeyPress(KeyAscii As Integer)
Event FrmKeyUp(KeyCode As Integer, Shift As Integer)
Event TreLOVDblClick()
Event TreLOVExpand(ByVal Node As MSComctlLib.Node)
Event TxtLovChange()
Event TxtLovKeyPress(KeyAscii As Integer)

Public Sist$, Ver$, Cia$
'Public IdField, Id, Cab
 
'* Variaveis de Impressão
Public vetCab
Const FONTE_TIT = "Times New Roman"
Const FONTE_VAL = "Times New Roman"
Const TAM_FONTE_TIT = 14
Const TAM_FONTE_VAL = 9
Private Sub CmdCancel_Click()
   RaiseEvent CmdCancelClick
End Sub
Private Sub CmdFiltrar_Click()
   RaiseEvent CmdFiltrarClick
End Sub
Private Sub CmdImprimir_Click()
   RaiseEvent CmdImprimirClick
End Sub
Private Sub CmdOk_Click()
   RaiseEvent CmdOkClick
End Sub
Private Sub CmdToExcel_Click()
   RaiseEvent CmdToExcelClick
End Sub
Private Sub CmdGrid_Click()
   RaiseEvent CmdGridClick
End Sub
Private Sub CmdTreeView_Click()
   RaiseEvent CmdTreeViewClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Click()
   On Error Resume Next
   Me.GrdLov.SetFocus
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   RaiseEvent FrmKeyPress(KeyAscii)
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent FrmKeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   RaiseEvent QueryUnload(Cancel, UnloadMode)
End Sub
Private Sub Form_Resize()
   RaiseEvent Resize
End Sub
Private Sub Form_Unload(Cancel As Integer)
   RaiseEvent Unload(Cancel)
End Sub
Private Sub GrdLov_Click()
   RaiseEvent GrdClick
End Sub
Private Sub GrdLov_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
   RaiseEvent GrdCompare(Row1, Row2, Cmp)
End Sub
Private Sub GrdLov_DblClick()
   RaiseEvent GrdDblClick  '* CmdOper(0)
End Sub
Private Sub GrdLov_KeyPress(KeyAscii As Integer)
   Dim iCol As Integer, iRow As Integer
   Dim Key As Integer
   iCol = Val(Chr(KeyAscii))
   Key = Asc(UCase(Chr(KeyAscii)))
   
   If Key < vbKey0 Or Key > vbKeyZ Then Exit Sub
   
   Me.TxtLov.Text = Me.TxtLov.Text + UCase(Chr(KeyAscii))
   Me.TxtLov.SelStart = Len(Me.TxtLov.Text)
   Me.TxtLov.SelLength = 1
End Sub
Private Sub GrdLov_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeySpace Then
      Me.TxtLov.Text = Me.TxtLov.Text + " "
   End If
End Sub
Private Sub GrdLov_LeaveCell()
   RaiseEvent GrdLovLeaveCell
End Sub
Private Sub GrdLov_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim ObjImg As Object
   If y <= Me.GrdLov.RowHeight(0) Then
      Me.GrdLov.MousePointer = flexCustom
   Else
      Me.GrdLov.MousePointer = flexDefault
   End If
End Sub
Private Sub GrdLov_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'   Dim Sql$
'   Me.TxtLov.Text = ""
'   With Me.GrdLov
'      If .MouseRow = 0 Then
'         .Visible = False
'         Call OrdenarMSGrid(Me.DataLov, Me.GrdLov, .MouseCol)
'         .Visible = True
'         .SetFocus
'      End If
'   End With
End Sub
Private Sub GrdLov_RowColChange()
'   RaiseEvent GrdLovLeaveCell
End Sub
Private Sub GrdLov_SelChange()
   RaiseEvent GrdLovSelChange
End Sub
Private Sub LblLov_Click(index As Integer)
   Select Case index
      Case 0
      Case 1
      Case 2
      Case 3
         Me.TxtLov.Text = Me.TxtLov.Tag
    End Select
End Sub
Private Sub TreLOV_DblClick()
   RaiseEvent TreLOVDblClick
End Sub
Private Sub TreLOV_Expand(ByVal Node As MSComctlLib.Node)
   RaiseEvent TreLOVExpand(Node)
End Sub

Private Sub TreLOV_KeyPress(KeyAscii As Integer)
   Me.TxtLov.Text = Me.TxtLov.Text + UCase(Chr(KeyAscii))
   Me.TxtLov.SelStart = Len(Me.TxtLov.Text)
   Me.TxtLov.SelLength = 1
End Sub

Private Sub TxtLov_Change()
   RaiseEvent TxtLovChange
End Sub
Private Sub TxtLov_GotFocus()
   'TxtLov.SelStart = 0
   'TxtLov.SelLength = Len(TxtLov)
End Sub
Private Sub TxtLov_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtLovKeyPress(KeyAscii)
'   Call  MSGrdProcurar(Me.DataLov, Me.GrdLov, Me.TxtLov, KeyAscii)
'   Me.TxtLov.SetFocus
End Sub
