VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#11.2#0"; "Codejock.ReportControl.v11.2.2.ocx"
Begin VB.Form FrmLov2 
   AutoRedraw      =   -1  'True
   Caption         =   "GrdLov"
   ClientHeight    =   6735
   ClientLeft      =   2760
   ClientTop       =   510
   ClientWidth     =   5100
   Icon            =   "Lov2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6735
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin XtremeReportControl.ReportControl GrdLov 
      Height          =   5295
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   4815
      _Version        =   720898
      _ExtentX        =   8493
      _ExtentY        =   9340
      _StockProps     =   64
   End
   Begin Threed.SSCommand CmdToExcel 
      Height          =   330
      Left            =   1440
      TabIndex        =   9
      ToolTipText     =   "Exportar Para o Excell"
      Top             =   0
      Width           =   330
      _Version        =   65536
      _ExtentX        =   582
      _ExtentY        =   582
      _StockProps     =   78
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      MouseIcon       =   "Lov2.frx":0442
      Picture         =   "Lov2.frx":045E
   End
   Begin Threed.SSCommand CmdFiltrar 
      Height          =   330
      Left            =   4500
      TabIndex        =   6
      ToolTipText     =   "Filtrar Lista"
      Top             =   5880
      Width           =   330
      _Version        =   65536
      _ExtentX        =   582
      _ExtentY        =   582
      _StockProps     =   78
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Font3D          =   3
      RoundedCorners  =   0   'False
      MouseIcon       =   "Lov2.frx":047A
      Picture         =   "Lov2.frx":0496
   End
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
      TabIndex        =   1
      Top             =   5880
      Width           =   3375
   End
   Begin Threed.SSCommand CmdImprimir 
      Height          =   330
      Left            =   1800
      TabIndex        =   0
      ToolTipText     =   "Visualizar Impressão"
      Top             =   0
      Width           =   330
      _Version        =   65536
      _ExtentX        =   582
      _ExtentY        =   582
      _StockProps     =   78
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "Lov2.frx":04B2
   End
   Begin Threed.SSCommand CmdLovOper 
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   3
      Top             =   6300
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&OK"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSCommand CmdLovOper 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   6300
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Cancela"
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin MSComctlLib.TreeView TreLOV 
      Height          =   5355
      Left            =   120
      TabIndex        =   5
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
   Begin Threed.SSCommand CmdTreeView 
      Height          =   330
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Exibir Estrutura"
      Top             =   0
      Width           =   330
      _Version        =   65536
      _ExtentX        =   582
      _ExtentY        =   582
      _StockProps     =   78
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "Lov2.frx":04CE
   End
   Begin Threed.SSCommand CmdGrid 
      Height          =   330
      Left            =   480
      TabIndex        =   8
      ToolTipText     =   "Exibir Lista"
      Top             =   0
      Width           =   330
      _Version        =   65536
      _ExtentX        =   582
      _ExtentY        =   582
      _StockProps     =   78
      BevelWidth      =   1
      RoundedCorners  =   0   'False
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
      TabIndex        =   2
      ToolTipText     =   "* Use [Esc] para limpar o campo ""Localizar"""
      Top             =   5880
      Width           =   1050
   End
End
Attribute VB_Name = "FrmLov2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event Activate()
Event Load()
Event Resize()
Event QueryUnload(Cancel, UnloadMode)
Event Unload(Cancel As Integer)
Event CmdOperClick(Index As Integer)
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

Private Sub CmdFiltrar_Click()
   RaiseEvent CmdFiltrarClick
End Sub
Private Sub CmdLovOper_Click(Index As Integer)
   RaiseEvent CmdOperClick(Index)
End Sub
Private Sub CmdImprimir_Click()
   RaiseEvent CmdImprimirClick
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

Private Sub GrdLov_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
   Dim ObjImg As Object
   
'   If y <= Me.GrdLov.RowHeight(0) Then
'      Me.GrdLov.MousePointer = flexCustom
'   Else
'      Me.GrdLov.MousePointer = flexDefault
'   End If

End Sub
Private Sub GrdLov_SelChange()
   RaiseEvent GrdLovSelChange
End Sub
Private Sub LblLov_Click(Index As Integer)
   Select Case Index
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

